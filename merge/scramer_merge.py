from __future__ import annotations

import datetime as dt
import json
import os
import re
import sys
from pathlib import Path
from typing import Any, Mapping

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from merge.Ashare_merge import build_payload as build_ashare_payload
from merge.Hshare_merge import fetch as fetch_hshare
from crawlers.convertible_bond import fetch as fetch_bond
from utils import get_calendar_range, init_calendar_data, normalize_fetch_output


REFERENCE_DATE = dt.datetime.now().date().strftime("%Y-%m-%d")
DATA_DIR = ROOT_DIR / "data"
OUTPUT_FILE = DATA_DIR / "calendar_data.json"
ASHARE_FILE = DATA_DIR / "Asharecalendar_data.json"
BOND_FILE = DATA_DIR / "bondcalendar_data.json"
HSHARE_FILE = DATA_DIR / "Hsharecalendar_data.json"
DATE_KEY_RE = re.compile(r"\d{4}-\d{2}-\d{2}")
R2_ENDPOINT_URL = "https://0cb4978c638c310279a1d85ef69f1d23.r2.cloudflarestorage.com"
R2_ENV_VARS = (
    "R2_ACCESS_KEY_ID",
    "R2_SECRET_ACCESS_KEY",
    "R2_BUCKET_NAME",
)
UNLOCK_DB_FILE = ROOT_DIR / "frontend" / "unlock_db.js"
UNLOCK_DB_RE = re.compile(r"window\.UNLOCK_DB\s*=\s*(\[[\s\S]*\])\s*;?\s*$")
UNLOCK_MARKETS = {"SH", "SZ"}
REITS_CODE_PREFIXES = ("18", "50")


def _empty_map(reference_date: str) -> dict[str, dict[str, list[Any]]]:
    date_list = get_calendar_range(reference_date)
    return normalize_fetch_output(
        init_calendar_data(date_list),
        date_list=date_list,
        reference_date=reference_date,
    )


def _collect_date_list(reference_date: str, *parts: Any) -> list[str]:
    keys = set(get_calendar_range(reference_date))
    for part in parts:
        if isinstance(part, Mapping):
            for key in part.keys():
                if isinstance(key, str) and DATE_KEY_RE.fullmatch(key):
                    keys.add(key)
        elif isinstance(part, list):
            for key in part:
                if isinstance(key, str) and DATE_KEY_RE.fullmatch(key):
                    keys.add(key)
    return sorted(keys)


def _safe_build_ashare(reference_date: str) -> dict[str, Any]:
    try:
        return build_ashare_payload(reference_date=reference_date)
    except Exception as exc:
        print(f"[SCRAMER][ASHARE] failed: {exc}")
        date_list = get_calendar_range(reference_date)
        empty_map = _empty_map(reference_date)
        return {
            "generated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "reference_date": reference_date,
            "date_list": date_list,
            "a_share": {"all": empty_map, "hs": empty_map, "bj": empty_map},
            "all": empty_map,
            "hs": empty_map,
            "bj": empty_map,
            "bond": {},
            "h_share": {},
        }


def _safe_fetch_bond(reference_date: str) -> dict[str, dict[str, list[Any]]]:
    try:
        return fetch_bond(reference_date=reference_date, verbose=False)
    except Exception as exc:
        print(f"[SCRAMER][BOND] failed: {exc}")
        return _empty_map(reference_date)


def _safe_fetch_hshare(reference_date: str) -> dict[str, dict[str, list[Any]]]:
    try:
        return fetch_hshare(reference_date=reference_date, verbose=False)
    except Exception as exc:
        print(f"[SCRAMER][HSHARE] failed: {exc}")
        return _empty_map(reference_date)


def _write_json(path: str | Path, payload: dict[str, Any]) -> None:
    path_obj = Path(path)
    path_obj.parent.mkdir(parents=True, exist_ok=True)
    path_obj.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _is_reits_code(code: str) -> bool:
    return code.startswith(REITS_CODE_PREFIXES)


def _parse_unlock_db(path: str | Path) -> list[dict[str, Any]]:
    path_obj = Path(path)
    if not path_obj.exists():
        return []

    raw = path_obj.read_text(encoding="utf-8")
    match = UNLOCK_DB_RE.search(raw)
    if not match:
        raise RuntimeError(f"unlock db parse failed: {path_obj}")

    array_literal = match.group(1)
    json_text = re.sub(
        r'([{,]\s*)([A-Za-z_][A-Za-z0-9_]*)\s*:',
        r'\1"\2":',
        array_literal,
    )
    json_text = re.sub(r",(\s*[\]}])", r"\1", json_text)
    records = json.loads(json_text)
    if not isinstance(records, list):
        raise RuntimeError(f"unlock db is not a list: {path_obj}")

    normalized: list[dict[str, Any]] = []
    for item in records:
        if not isinstance(item, Mapping):
            continue
        normalized.append(
            {
                "code": str(item.get("code", "")).strip(),
                "name": str(item.get("name", "")).strip(),
                "market": str(item.get("market", "")).strip().upper(),
                "listing_date": item.get("listing_date"),
                "lock_months": item.get("lock_months"),
                "lock_day": item.get("lock_day"),
            }
        )
    return normalized


def _format_js_nullable_string(value: Any) -> str:
    if value is None:
        return "null"
    text = str(value).strip()
    return "null" if not text else json.dumps(text, ensure_ascii=False)


def _format_js_nullable_number(value: Any) -> str:
    if value in (None, ""):
        return "null"
    try:
        number = int(value)
    except (TypeError, ValueError):
        return "null"
    return str(number)


def _sort_unlock_records(records: list[dict[str, Any]]) -> None:
    records.sort(
        key=lambda item: (
            str(item.get("listing_date") or "9999-99-99"),
            str(item.get("code") or ""),
        )
    )


def _write_unlock_db(path: str | Path, records: list[dict[str, Any]]) -> None:
    path_obj = Path(path)
    path_obj.parent.mkdir(parents=True, exist_ok=True)
    lines = ["window.UNLOCK_DB = ["]
    for index, record in enumerate(records):
        suffix = "," if index < len(records) - 1 else ""
        lines.append(
            "  { code: %s, name: %s, market: %s, listing_date: %s, lock_months: %s, lock_day: %s }%s"
            % (
                json.dumps(str(record.get("code", "")).strip(), ensure_ascii=False),
                json.dumps(str(record.get("name", "")).strip(), ensure_ascii=False),
                json.dumps(str(record.get("market", "")).strip().upper(), ensure_ascii=False),
                _format_js_nullable_string(record.get("listing_date")),
                _format_js_nullable_number(record.get("lock_months")),
                _format_js_nullable_number(record.get("lock_day")),
                suffix,
            )
        )
    lines.append("];")
    path_obj.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _sync_unlock_db_from_ashare(source_data: Mapping[str, Any], path: str | Path = UNLOCK_DB_FILE) -> None:
    records = _parse_unlock_db(path)
    seen_codes = {
        str(record.get("code", "")).strip()
        for record in records
        if str(record.get("code", "")).strip()
    }
    added = 0

    for date_key in sorted(source_data):
        if not DATE_KEY_RE.fullmatch(date_key):
            continue
        source_day = source_data.get(date_key)
        if not isinstance(source_day, Mapping):
            continue
        for item in source_day.get("listing", []):
            if not isinstance(item, Mapping):
                continue
            code = str(item.get("code", "")).strip()
            name = str(item.get("name", "")).strip()
            market = str(item.get("market", "")).strip().upper()
            if not code or code in seen_codes or market not in UNLOCK_MARKETS:
                continue
            records.append(
                {
                    "code": code,
                    "name": name or code,
                    "market": market,
                    "listing_date": None,
                    "lock_months": None if _is_reits_code(code) else 6,
                    "lock_day": None,
                }
            )
            seen_codes.add(code)
            added += 1

    if added == 0:
        print("[SCRAMER][UNLOCK_DB] skipped: no new SH/SZ listings")
        return

    _sort_unlock_records(records)
    _write_unlock_db(path, records)
    print(f"[SCRAMER][UNLOCK_DB] written: {path} (+{added})")


def _get_r2_config() -> dict[str, str] | None:
    config = {name: os.getenv(name, "").strip() for name in R2_ENV_VARS}
    missing = [name for name, value in config.items() if not value]
    if len(missing) == len(R2_ENV_VARS):
        return None
    if missing:
        raise RuntimeError(
            f"missing R2 environment variables: {', '.join(missing)}"
        )
    return config


def _upload_json_files_to_r2(*paths: str | Path) -> None:
    config = _get_r2_config()
    if not config:
        print("[SCRAMER][R2] skipped: R2 environment variables are not set")
        return

    import boto3

    client = boto3.client(
        "s3",
        endpoint_url=R2_ENDPOINT_URL,
        aws_access_key_id=config["R2_ACCESS_KEY_ID"],
        aws_secret_access_key=config["R2_SECRET_ACCESS_KEY"],
        region_name="auto",
    )

    for path in paths:
        path_obj = Path(path)
        client.upload_file(
            str(path_obj),
            config["R2_BUCKET_NAME"],
            path_obj.name,
        )
        print(f"已上传: {path_obj.name}")


def build_payload(reference_date: str = REFERENCE_DATE) -> dict[str, Any]:
    ashare_payload = _safe_build_ashare(reference_date)
    bond_map = _safe_fetch_bond(reference_date)
    hshare_map = _safe_fetch_hshare(reference_date)
    date_list = _collect_date_list(
        reference_date,
        ashare_payload.get("date_list", []),
        ashare_payload.get("all", {}),
        ashare_payload.get("hs", {}),
        ashare_payload.get("bj", {}),
        bond_map,
        hshare_map,
    )
    bond_map = normalize_fetch_output(
        bond_map,
        date_list=date_list,
        reference_date=reference_date,
    )
    hshare_map = normalize_fetch_output(
        hshare_map,
        date_list=date_list,
        reference_date=reference_date,
    )
    generated_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    bond_payload = {
        "generated_at": generated_at,
        "reference_date": reference_date,
        "date_list": date_list,
        "bond": bond_map,
        "convertible_bond": bond_map,
    }
    hshare_payload = {
        "generated_at": generated_at,
        "reference_date": reference_date,
        "date_list": date_list,
        "h_share": hshare_map,
        "hk": hshare_map,
    }

    _write_json(ASHARE_FILE, ashare_payload)
    _write_json(BOND_FILE, bond_payload)
    _write_json(HSHARE_FILE, hshare_payload)

    return {
        "generated_at": generated_at,
        "reference_date": reference_date,
        "date_list": date_list,
        "a_share": ashare_payload.get("a_share", {}),
        "all": ashare_payload.get("all", {}),
        "hs": ashare_payload.get("hs", {}),
        "bj": ashare_payload.get("bj", {}),
        "bond": bond_map,
        "convertible_bond": bond_map,
        "h_share": hshare_map,
        "hk": hshare_map,
    }


def main() -> None:
    payload = build_payload(reference_date=REFERENCE_DATE)
    _write_json(OUTPUT_FILE, payload)
    _sync_unlock_db_from_ashare((payload.get("a_share") or {}).get("all") or {})
    print(f"[SCRAMER] written: {OUTPUT_FILE}")
    print(f"[SCRAMER] written: {ASHARE_FILE}, {BOND_FILE}, {HSHARE_FILE}")
    _upload_json_files_to_r2(
        ASHARE_FILE,
        BOND_FILE,
        HSHARE_FILE,
        OUTPUT_FILE,
    )


if __name__ == "__main__":
    main()
