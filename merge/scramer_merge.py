from __future__ import annotations

import datetime as dt
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Any, Mapping

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from merge.Ashare_merge import build_payload as build_ashare_payload
from merge.Hshare_merge import fetch as fetch_hshare
from crawlers.convertible_bond import fetch as fetch_bond
from utils import EMPTY_DAY, get_calendar_range, init_calendar_data, normalize_fetch_output


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
ASHARE_STABILIZE_ATTEMPTS = max(1, int(os.getenv("SCRAMER_ASHARE_ATTEMPTS", "3")))
ASHARE_STABILIZE_SLEEP_SECONDS = max(0.0, float(os.getenv("SCRAMER_ASHARE_SLEEP_SECONDS", "6")))


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
    payloads: list[dict[str, Any]] = []

    for attempt in range(ASHARE_STABILIZE_ATTEMPTS):
        try:
            payload = build_ashare_payload(reference_date=reference_date)
            payloads.append(payload)
            print(
                f"[SCRAMER][ASHARE] captured attempt "
                f"{attempt + 1}/{ASHARE_STABILIZE_ATTEMPTS}"
            )
        except Exception as exc:
            print(f"[SCRAMER][ASHARE] attempt {attempt + 1} failed: {exc}")

        if attempt + 1 < ASHARE_STABILIZE_ATTEMPTS and ASHARE_STABILIZE_SLEEP_SECONDS > 0:
            print(
                f"[SCRAMER][ASHARE] waiting "
                f"{ASHARE_STABILIZE_SLEEP_SECONDS:.1f}s before next capture"
            )
            time.sleep(ASHARE_STABILIZE_SLEEP_SECONDS)

    if payloads:
        return _stabilize_ashare_payload(reference_date, payloads)

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


def _item_identity(item: Any) -> str:
    if isinstance(item, Mapping):
        code = str(item.get("code", "")).strip()
        name = str(item.get("name", "")).strip()
        market = str(item.get("market", "")).strip().upper()
        if code:
            return f"code:{code}|{market}"
        if name:
            return f"name:{name}|{market}"
        return "dict:" + json.dumps(item, ensure_ascii=False, sort_keys=True)
    return "scalar:" + str(item)


def _item_richness(item: Any) -> int:
    if not isinstance(item, Mapping):
        return 0
    return sum(1 for key in ("code", "name", "market") if str(item.get(key, "")).strip())


def _merge_calendar_maps_prefer_earliest(
    reference_date: str,
    *maps: Mapping[str, Any],
) -> dict[str, dict[str, list[Any]]]:
    date_list = _collect_date_list(reference_date, *maps)
    merged = init_calendar_data(date_list)
    picked: dict[str, dict[str, tuple[str, Any]]] = {
        event_key: {} for event_key in EMPTY_DAY.keys()
    }

    for source in maps:
        if not isinstance(source, Mapping):
            continue
        for date_key, day_data in source.items():
            if not (isinstance(date_key, str) and DATE_KEY_RE.fullmatch(date_key)):
                continue
            if not isinstance(day_data, Mapping):
                continue

            for event_key in EMPTY_DAY.keys():
                raw_items = day_data.get(event_key, [])
                if not isinstance(raw_items, list):
                    raw_items = [raw_items]

                for item in raw_items:
                    identity = _item_identity(item)
                    existing = picked[event_key].get(identity)
                    if existing is None:
                        picked[event_key][identity] = (date_key, item)
                        continue

                    existing_date, existing_item = existing
                    if date_key < existing_date:
                        picked[event_key][identity] = (date_key, item)
                        continue
                    if date_key == existing_date and _item_richness(item) > _item_richness(existing_item):
                        picked[event_key][identity] = (date_key, item)

    for event_key, values in picked.items():
        ordered = sorted(values.items(), key=lambda entry: (entry[1][0], entry[0]))
        for _, (date_key, item) in ordered:
            if date_key in merged:
                merged[date_key][event_key].append(item)

    return normalize_fetch_output(
        merged,
        date_list=date_list,
        reference_date=reference_date,
    )


def _stabilize_ashare_payload(
    reference_date: str,
    payloads: list[dict[str, Any]],
) -> dict[str, Any]:
    all_maps = [payload.get("all", {}) for payload in payloads if isinstance(payload, Mapping)]
    hs_maps = [payload.get("hs", {}) for payload in payloads if isinstance(payload, Mapping)]
    bj_maps = [payload.get("bj", {}) for payload in payloads if isinstance(payload, Mapping)]
    date_sources: list[Any] = []

    for payload in payloads:
        if not isinstance(payload, Mapping):
            continue
        date_sources.append(payload.get("date_list", []))
        date_sources.append(payload.get("all", {}))
        date_sources.append(payload.get("hs", {}))
        date_sources.append(payload.get("bj", {}))

    date_list = _collect_date_list(reference_date, *date_sources)
    all_map = _merge_calendar_maps_prefer_earliest(reference_date, *all_maps)
    hs_map = _merge_calendar_maps_prefer_earliest(reference_date, *hs_maps)
    bj_map = _merge_calendar_maps_prefer_earliest(reference_date, *bj_maps)
    generated_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    return {
        "generated_at": generated_at,
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
