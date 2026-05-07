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
R2_ENV_VARS = (
    "R2_ACCESS_KEY_ID",
    "R2_SECRET_ACCESS_KEY",
    "R2_ACCOUNT_ID",
    "R2_BUCKET_NAME",
)


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


def _get_r2_config() -> dict[str, str] | None:
    config = {name: os.getenv(name, "").strip() for name in R2_ENV_VARS}
    missing = [name for name, value in config.items() if not value]
    if len(missing) == len(R2_ENV_VARS):
        return None
    if missing:
        raise RuntimeError(
            f"missing R2 environment variables: {', '.join(missing)}"
        )
    config["endpoint_url"] = (
        f"https://{config['R2_ACCOUNT_ID']}.r2.cloudflarestorage.com"
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
        endpoint_url=config["endpoint_url"],
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
            ExtraArgs={
                "ACL": "public-read",
                "ContentType": "application/json",
            },
        )
        print(f"[SCRAMER][R2] uploaded: {path_obj.name}")


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
