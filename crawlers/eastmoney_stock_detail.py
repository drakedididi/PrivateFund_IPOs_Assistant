from __future__ import annotations

import argparse
import json
import re
import sys
import time
from decimal import Decimal, InvalidOperation
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Iterable, Mapping

import requests


DETAIL_URL_TEMPLATE = "https://data.eastmoney.com/xg/xg/detail/{code}.html"
OTHER_DETAILS_API = "https://datacenter-web.eastmoney.com/api/data/v1/get"
STOCK_CODE_RE = re.compile(r"\d{6}")
INFO_ASSIGNMENT_RE = re.compile(r"\bvar\s+info\s*=\s*")
DEFAULT_TIMEOUT_SECONDS = 30
DEFAULT_RETRIES = 3
DEFAULT_HEADERS = {
    "Accept": "text/html,application/xhtml+xml,application/json;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.7",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
}


class StockDetailFetchError(RuntimeError):
    pass


class _ScriptCollector(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._in_script = False
        self._chunks: list[str] = []
        self.scripts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() == "script":
            self._in_script = True
            self._chunks = []

    def handle_data(self, data: str) -> None:
        if self._in_script:
            self._chunks.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() != "script" or not self._in_script:
            return
        self.scripts.append("".join(self._chunks))
        self._in_script = False
        self._chunks = []


def _normalize_code(value: Any) -> str:
    code = str(value or "").strip()
    if not STOCK_CODE_RE.fullmatch(code):
        raise ValueError(f"invalid stock code: {value!r}")
    return code


def _extract_info_payload(html: str) -> dict[str, Any]:
    parser = _ScriptCollector()
    parser.feed(html)
    decoder = json.JSONDecoder()

    for script in parser.scripts:
        assignment = INFO_ASSIGNMENT_RE.search(script)
        if not assignment:
            continue
        object_start = script.find("{", assignment.end())
        if object_start < 0:
            continue
        try:
            payload, _ = decoder.raw_decode(script[object_start:])
        except json.JSONDecodeError:
            continue
        if isinstance(payload, dict):
            return payload

    raise StockDetailFetchError("Eastmoney detail page does not contain a valid info payload")


def _to_share_count(value: Any, scale: int = 1) -> int | None:
    if value in (None, "", "-"):
        return None
    try:
        number = Decimal(str(value)) * scale
    except (InvalidOperation, ValueError):
        return None
    if not number.is_finite() or number < 0 or number != number.to_integral_value():
        return None
    return int(number)


def _infer_market(code: str, trade_market: Any = "") -> str:
    market_text = str(trade_market or "").strip()
    if "上海" in market_text:
        return "SH"
    if "深圳" in market_text:
        return "SZ"
    if code.startswith("6"):
        return "SH"
    if code.startswith(("0", "3")):
        return "SZ"
    return ""


def _get_with_retries(
    session: requests.Session,
    url: str,
    *,
    timeout: int,
    retries: int,
    params: Mapping[str, Any] | None = None,
    headers: Mapping[str, str] | None = None,
) -> requests.Response:
    last_error: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            response = session.get(
                url,
                params=params,
                headers=dict(headers or {}),
                timeout=timeout,
            )
            response.raise_for_status()
            return response
        except requests.RequestException as exc:
            last_error = exc
            if attempt < retries:
                time.sleep(min(0.5 * attempt, 2.0))
    raise StockDetailFetchError(f"request failed after {retries} attempts: {url}: {last_error}")


def fetch_stock_detail(
    code: str,
    *,
    session: requests.Session | None = None,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
    retries: int = DEFAULT_RETRIES,
) -> dict[str, Any]:
    normalized_code = _normalize_code(code)
    own_session = session is None
    client = session or requests.Session()
    client.headers.update(DEFAULT_HEADERS)
    detail_url = DETAIL_URL_TEMPLATE.format(code=normalized_code)

    try:
        page_response = _get_with_retries(
            client,
            detail_url,
            timeout=timeout,
            retries=retries,
        )
        page_response.encoding = page_response.apparent_encoding or "utf-8"
        info = _extract_info_payload(page_response.text)

        payload_code = _normalize_code(info.get("SECURITY_CODE"))
        if payload_code != normalized_code:
            raise StockDetailFetchError(
                f"detail payload code mismatch: requested {normalized_code}, received {payload_code}"
            )

        industry_response = _get_with_retries(
            client,
            OTHER_DETAILS_API,
            timeout=timeout,
            retries=retries,
            params={
                "reportName": "RPT_IPO_OTHERDETAILS",
                "columns": "INDUSTRY",
                "filter": f'(SECURITY_CODE="{normalized_code}")',
                "source": "WEB",
                "client": "WEB",
            },
            headers={"Referer": detail_url},
        )
        try:
            industry_payload = industry_response.json()
        except requests.JSONDecodeError as exc:
            raise StockDetailFetchError("Eastmoney industry endpoint returned invalid JSON") from exc

        industry_rows = ((industry_payload.get("result") or {}).get("data") or [])
        industry = ""
        if industry_rows and isinstance(industry_rows[0], Mapping):
            industry = str(industry_rows[0].get("INDUSTRY") or "").strip()

        return {
            "code": normalized_code,
            "name": str(info.get("SECURITY_NAME") or "").strip(),
            "market": _infer_market(normalized_code, info.get("TRADE_MARKET")),
            "industry": industry or None,
            "offline_placing_num": _to_share_count(info.get("OFFLINE_PLACING_NUM")),
            # The detail page formats ISSUE_NUM with zoom: 4 before showing shares.
            "issue_num": _to_share_count(info.get("ISSUE_NUM"), scale=10_000),
            "source_url": detail_url,
        }
    finally:
        if own_session:
            client.close()


def fetch_stock_details(
    codes: Iterable[str],
    *,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
    retries: int = DEFAULT_RETRIES,
) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    with requests.Session() as session:
        session.headers.update(DEFAULT_HEADERS)
        for code in codes:
            results.append(
                fetch_stock_detail(
                    code,
                    session=session,
                    timeout=timeout,
                    retries=retries,
                )
            )
    return results


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Fetch Eastmoney IPO industry and issuance details by stock code.",
    )
    parser.add_argument("codes", nargs="+", help="One or more six-digit Shanghai/Shenzhen stock codes")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT_SECONDS)
    parser.add_argument("--retries", type=int, default=DEFAULT_RETRIES)
    parser.add_argument("--output", type=Path, help="Optional UTF-8 JSON output path")
    return parser


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    args = _build_parser().parse_args()
    try:
        payload = fetch_stock_details(
            args.codes,
            timeout=max(1, args.timeout),
            retries=max(1, args.retries),
        )
    except (StockDetailFetchError, ValueError) as exc:
        print(f"[EASTMONEY][DETAIL] failed: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc

    output = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output + "\n", encoding="utf-8")
        print(f"[EASTMONEY][DETAIL] written: {args.output}")
    else:
        print(output)


if __name__ == "__main__":
    main()
