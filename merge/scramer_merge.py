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
TRADING_HOLIDAYS_FILE = ROOT_DIR / "data" / "trading_holidays.json"
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


def _load_trading_holiday_keys(path: str | Path = TRADING_HOLIDAYS_FILE) -> set[str]:
    path_obj = Path(path)
    payload = json.loads(path_obj.read_text(encoding="utf-8"))
    values = payload if isinstance(payload, list) else payload.get("vacation_dates", [])
    if not isinstance(values, list):
        raise RuntimeError(f"invalid trading holidays payload: {path_obj}")

    return {
        value.strip()
        for value in values
        if isinstance(value, str) and DATE_KEY_RE.fullmatch(value.strip())
    }


def _parse_non_negative_int(value: Any) -> int | None:
    if value in (None, ""):
        return None
    try:
        number = int(value)
    except (TypeError, ValueError):
        return None
    return number if number >= 0 else None


def _add_months_safe(base_date: dt.date, months: int) -> dt.date:
    target_month = base_date.month - 1 + months
    year = base_date.year + target_month // 12
    month = target_month % 12 + 1
    month_start = dt.date(year, month, 1)
    if month == 12:
        next_month_start = dt.date(year + 1, 1, 1)
    else:
        next_month_start = dt.date(year, month + 1, 1)
    max_day = (next_month_start - month_start).days
    return dt.date(year, month, min(base_date.day, max_day))


def _is_trading_date(value: dt.date, holiday_keys: set[str]) -> bool:
    if value.weekday() >= 5:
        return False
    return value.strftime("%Y-%m-%d") not in holiday_keys


def _calculate_unlock_trade_date(
    listing_date: Any,
    lock_months: Any,
    lock_day: Any,
    holiday_keys: set[str],
) -> dt.date | None:
    text = str(listing_date or "").strip()
    if not DATE_KEY_RE.fullmatch(text):
        return None
    try:
        base_date = dt.datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError:
        return None

    day_num = _parse_non_negative_int(lock_day)
    if day_num is not None:
        trade_date = base_date + dt.timedelta(days=day_num)
    else:
        month_num = _parse_non_negative_int(lock_months)
        if month_num is None or month_num <= 0:
            return None
        trade_date = _add_months_safe(base_date, month_num)

    while not _is_trading_date(trade_date, holiday_keys):
        trade_date += dt.timedelta(days=1)
    return trade_date


def _purge_expired_unlock_records(
    records: list[dict[str, Any]],
    reference_date: str,
    holiday_keys: set[str],
) -> int:
    week_start = dt.datetime.strptime(reference_date, "%Y-%m-%d").date()
    week_start -= dt.timedelta(days=week_start.weekday())

    kept_records: list[dict[str, Any]] = []
    removed = 0
    for record in records:
        unlock_date = _calculate_unlock_trade_date(
            record.get("listing_date"),
            record.get("lock_months"),
            record.get("lock_day"),
            holiday_keys,
        )
        if unlock_date and unlock_date < week_start:
            removed += 1
            continue
        kept_records.append(record)

    if removed:
        records[:] = kept_records
    return removed


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


def _sync_unlock_db_from_calendar(
    source_data: Mapping[str, Any],
    path: str | Path = UNLOCK_DB_FILE,
    reference_date: str = REFERENCE_DATE,
) -> None:
    records = _parse_unlock_db(path)
    record_by_code: dict[str, dict[str, Any]] = {}
    for record in records:
        code = str(record.get("code", "")).strip()
        if code and code not in record_by_code:
            record_by_code[code] = record

    added = 0
    updated = 0
    holiday_keys = _load_trading_holiday_keys()

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
            if not code or market not in UNLOCK_MARKETS:
                continue

            existing = record_by_code.get(code)
            if existing:
                changed = False
                if not existing.get("listing_date"):
                    existing["listing_date"] = date_key
                    changed = True
                if not str(existing.get("name", "")).strip() and name:
                    existing["name"] = name
                    changed = True
                if not str(existing.get("market", "")).strip() and market:
                    existing["market"] = market
                    changed = True
                if changed:
                    updated += 1
                continue

            record = {
                "code": code,
                "name": name or code,
                "market": market,
                "listing_date": date_key,
                "lock_months": None if _is_reits_code(code) else 6,
                "lock_day": None,
            }
            records.append(record)
            record_by_code[code] = record
            added += 1

    removed = _purge_expired_unlock_records(records, reference_date=reference_date, holiday_keys=holiday_keys)

    if added == 0 and updated == 0 and removed == 0:
        print("[SCRAMER][UNLOCK_DB] skipped: no listing-driven updates")
        return

    _sort_unlock_records(records)
    _write_unlock_db(path, records)
    print(f"[SCRAMER][UNLOCK_DB] written: {path} (+{added}, ~{updated}, -{removed})")


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
    _sync_unlock_db_from_calendar(
        payload.get("all") or {},
        reference_date=str(payload.get("reference_date") or REFERENCE_DATE),
    )
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
