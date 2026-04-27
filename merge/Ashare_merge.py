from __future__ import annotations

import datetime as dt
import json
import re
import sys
from pathlib import Path
from typing import Any, Iterable, Mapping

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from crawlers.bse_crawler import fetch as fetch_bse
from crawlers.eastmoney_crawler import fetch as fetch_eastmoney
from crawlers.sse_crawler import fetch as fetch_sse
from crawlers.sz_crawler import fetch as fetch_sz
from utils import EMPTY_DAY, get_calendar_range, init_calendar_data, normalize_fetch_output


DATA_DIR = ROOT_DIR / "data"
OUTPUT_FILE = DATA_DIR / "Asharecalendar_data.json"
REFERENCE_DATE = dt.datetime.now().date().strftime("%Y-%m-%d")
DATE_KEY_RE = re.compile(r"\d{4}-\d{2}-\d{2}")


def _signature(item: Any) -> str:
    if isinstance(item, Mapping):
        return "dict:" + json.dumps(item, ensure_ascii=False, sort_keys=True)
    return "scalar:" + str(item)


def _iter_items(value: Any) -> Iterable[Any]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def _market_of(item: Any) -> str:
    if isinstance(item, Mapping):
        return str(item.get("market", "")).upper()
    return ""


def _allow_item(item: Any, allowed_markets: set[str] | None) -> bool:
    if allowed_markets is None:
        return True
    market = _market_of(item)
    if not market:
        return True
    return market in allowed_markets


def _merge_map(
    target: dict[str, dict[str, list[Any]]],
    source: Mapping[str, Mapping[str, Any]],
    allowed_markets: set[str] | None = None,
) -> None:
    for date_key, source_day in source.items():
        if date_key not in target:
            continue
        if not isinstance(source_day, Mapping):
            continue

        target_day = target[date_key]
        for event_key in EMPTY_DAY.keys():
            existing = target_day.get(event_key, [])
            existing_signatures = {_signature(x) for x in existing}

            for item in _iter_items(source_day.get(event_key, [])):
                if not _allow_item(item, allowed_markets):
                    continue
                sig = _signature(item)
                if sig in existing_signatures:
                    continue
                target_day[event_key].append(item)
                existing_signatures.add(sig)


def _safe_fetch(name: str, fn, reference_date: str) -> dict[str, dict[str, list[Any]]]:
    try:
        return fn(reference_date=reference_date, verbose=False)
    except Exception as exc:
        print(f"[MERGE][{name}] fetch failed: {exc}")
        return init_calendar_data(get_calendar_range(reference_date))


def _collect_date_list(
    reference_date: str,
    *sources: Mapping[str, Any],
) -> list[str]:
    keys = set(get_calendar_range(reference_date))
    for source in sources:
        if not isinstance(source, Mapping):
            continue
        for key in source.keys():
            if isinstance(key, str) and DATE_KEY_RE.fullmatch(key):
                keys.add(key)
    return sorted(keys)


def build_payload(reference_date: str = REFERENCE_DATE) -> dict[str, Any]:
    sse_data = _safe_fetch("SSE", fetch_sse, reference_date)
    sz_data = _safe_fetch("SZSE", fetch_sz, reference_date)
    bse_data = _safe_fetch("BSE", fetch_bse, reference_date)
    em_data = _safe_fetch("EASTMONEY", fetch_eastmoney, reference_date)
    date_list = _collect_date_list(reference_date, sse_data, sz_data, bse_data, em_data)

    hs_map = init_calendar_data(date_list)
    bj_map = init_calendar_data(date_list)
    all_map = init_calendar_data(date_list)

    hs_markets = {"SH", "SZ", "A"}
    bj_markets = {"BJ"}

    _merge_map(hs_map, sse_data, allowed_markets=hs_markets)
    _merge_map(hs_map, sz_data, allowed_markets=hs_markets)
    _merge_map(hs_map, em_data, allowed_markets=hs_markets)

    _merge_map(bj_map, bse_data, allowed_markets=bj_markets)
    _merge_map(bj_map, em_data, allowed_markets=bj_markets)

    _merge_map(all_map, sse_data, allowed_markets=None)
    _merge_map(all_map, sz_data, allowed_markets=None)
    _merge_map(all_map, bse_data, allowed_markets=None)
    _merge_map(all_map, em_data, allowed_markets=None)

    hs_map = normalize_fetch_output(hs_map, date_list=date_list, reference_date=reference_date)
    bj_map = normalize_fetch_output(bj_map, date_list=date_list, reference_date=reference_date)
    all_map = normalize_fetch_output(all_map, date_list=date_list, reference_date=reference_date)

    return {
        "generated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "reference_date": reference_date,
        "date_list": date_list,
        "a_share": {
            "all": all_map,
            "hs": hs_map,
            "bj": bj_map,
        },
        "all": all_map,
        "hs": hs_map,
        "bj": bj_map,
        "bond": {},
        "h_share": {},
    }


def main() -> None:
    payload = build_payload(reference_date=REFERENCE_DATE)
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_FILE.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"[MERGE] written: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
