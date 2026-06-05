from __future__ import annotations

import datetime as dt
import json
import os
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
from utils import EMPTY_DAY, clean_data, get_calendar_range, init_calendar_data, normalize_fetch_output


DATA_DIR = ROOT_DIR / "data"
OUTPUT_FILE = DATA_DIR / "Asharecalendar_data.json"
REFERENCE_DATE = dt.datetime.now().date().strftime("%Y-%m-%d")
DATE_KEY_RE = re.compile(r"\d{4}-\d{2}-\d{2}")
FETCH_ATTEMPTS = max(1, int(os.getenv("ASHARE_FETCH_ATTEMPTS", "3")))
FETCH_TIMEOUT_MS = max(30000, int(os.getenv("ASHARE_FETCH_TIMEOUT_MS", "60000")))


def _signature(item: Any) -> str:
    if isinstance(item, Mapping):
        return "dict:" + json.dumps(item, ensure_ascii=False, sort_keys=True)
    return "scalar:" + str(item)


def _normalize_identity_text(value: Any) -> str:
    text = clean_data(value)
    return re.sub(r"\s+", "", str(text)).strip().upper()


def _item_parts(item: Any) -> tuple[str, str, str]:
    if isinstance(item, Mapping):
        code = _normalize_identity_text(item.get("code", ""))
        name = _normalize_identity_text(item.get("name", ""))
        market = _normalize_identity_text(item.get("market", ""))
        return code, name, market
    return "", _normalize_identity_text(item), ""


def _same_item(left: Any, right: Any) -> bool:
    left_code, left_name, left_market = _item_parts(left)
    right_code, right_name, right_market = _item_parts(right)
    same_market = left_market in {"", "A"} or right_market in {"", "A"} or left_market == right_market
    if same_market and left_code and right_code and left_code == right_code:
        return True
    if same_market and left_name and right_name and left_name == right_name:
        return True
    return _signature(left) == _signature(right)


def _merge_item(existing: Any, new_item: Any) -> Any:
    if isinstance(existing, Mapping) and isinstance(new_item, Mapping):
        merged = dict(existing)
        for key, value in new_item.items():
            if merged.get(key) in (None, "") and value not in (None, ""):
                merged[key] = value
            elif key == "market" and str(merged.get(key, "")).upper() == "A" and str(value).upper() not in {"", "A"}:
                merged[key] = value
        return merged
    if isinstance(new_item, Mapping) and not isinstance(existing, Mapping):
        return dict(new_item)
    return existing


def _append_unique(items: list[Any], item: Any) -> None:
    for index, existing in enumerate(items):
        if _same_item(existing, item):
            items[index] = _merge_item(existing, item)
            return
    items.append(item)


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

            for item in _iter_items(source_day.get(event_key, [])):
                if not _allow_item(item, allowed_markets):
                    continue
                _append_unique(existing, item)


def _count_items(data: Mapping[str, Mapping[str, Any]]) -> dict[str, int]:
    counts = {event_key: 0 for event_key in EMPTY_DAY.keys()}
    for day_data in data.values():
        if not isinstance(day_data, Mapping):
            continue
        for event_key in EMPTY_DAY.keys():
            counts[event_key] += len(list(_iter_items(day_data.get(event_key, []))))
    return counts


def _format_counts(counts: Mapping[str, int]) -> str:
    return ", ".join(f"{key}={counts.get(key, 0)}" for key in EMPTY_DAY.keys())


def _merge_source_results(
    reference_date: str,
    sources: list[Mapping[str, Mapping[str, Any]]],
) -> dict[str, dict[str, list[Any]]]:
    date_list = _collect_date_list(reference_date, *sources)
    merged = init_calendar_data(date_list)
    for source in sources:
        _merge_map(merged, source, allowed_markets=None)
    return normalize_fetch_output(merged, date_list=date_list, reference_date=reference_date)


def _safe_fetch(
    name: str,
    fn,
    reference_date: str,
    attempts: int = FETCH_ATTEMPTS,
) -> dict[str, dict[str, list[Any]]]:
    results: list[Mapping[str, Mapping[str, Any]]] = []

    for attempt in range(1, attempts + 1):
        try:
            result = fn(
                reference_date=reference_date,
                timeout_ms=FETCH_TIMEOUT_MS,
                verbose=False,
            )
        except Exception as exc:
            print(f"[MERGE][{name}][attempt {attempt}/{attempts}] fetch failed: {exc}")
            continue

        results.append(result)
        print(
            f"[MERGE][{name}][attempt {attempt}/{attempts}] "
            f"{_format_counts(_count_items(result))}"
        )

    if not results:
        print(f"[MERGE][{name}] all attempts failed")
        return init_calendar_data(get_calendar_range(reference_date))

    merged = _merge_source_results(reference_date, results)
    print(f"[MERGE][{name}] merged {_format_counts(_count_items(merged))}")
    return merged


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
