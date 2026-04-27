from __future__ import annotations

import datetime as dt
import json
import re
import sys
from pathlib import Path
from typing import Any

from playwright.sync_api import Locator, sync_playwright

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from utils import clean_data, get_calendar_range, init_calendar_data, normalize_fetch_output


URL = "https://hk.eastmoney.com/ipolist.html"
TBODY_XPATH = '//*[@id="main"]/div/div[1]/div[2]/table/tbody'

REFERENCE_DATE = dt.datetime.now().date().strftime("%Y-%m-%d")
DATA_DIR = ROOT_DIR / "data"
OUTPUT_FILE = DATA_DIR / "Hsharecalendar_data.json"


def _is_valid_date_key(value: str) -> bool:
    return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", value))


def _extract_date_from_td(td: Locator, reference_date: str) -> str:
    span = td.locator("span")
    if span.count() > 0:
        parsed = clean_data(span.first.inner_text(), reference_date=reference_date)
        if isinstance(parsed, str) and _is_valid_date_key(parsed):
            return parsed

    parsed_text = clean_data(td.inner_text(), reference_date=reference_date)
    if isinstance(parsed_text, str) and _is_valid_date_key(parsed_text):
        return parsed_text
    return ""


def _extract_date_by_priority(cells: Locator, indexes: list[int], reference_date: str) -> str:
    for idx in indexes:
        if cells.count() <= idx:
            continue
        date_key = _extract_date_from_td(cells.nth(idx), reference_date=reference_date)
        if date_key:
            return date_key
    return ""


def _extract_code_name(cells: Locator, reference_date: str) -> tuple[str, str]:
    code = ""
    name = ""
    if cells.count() > 1:
        code = str(clean_data(cells.nth(1).inner_text(), reference_date=reference_date)).strip()
    if cells.count() > 2:
        name = str(clean_data(cells.nth(2).inner_text(), reference_date=reference_date)).strip()
    return code, name


def fetch(
    reference_date: str = REFERENCE_DATE,
    timeout_ms: int = 30000,
    verbose: bool = True,
) -> dict[str, dict[str, list[Any]]]:
    date_list = get_calendar_range(reference_date)
    date_set = set(date_list)
    raw_data = init_calendar_data(date_list)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto(URL, wait_until="domcontentloaded", timeout=timeout_ms)
        page.wait_for_selector(f"xpath={TBODY_XPATH}", timeout=timeout_ms)

        tbody = page.locator(f"xpath={TBODY_XPATH}")
        rows = tbody.first.locator("tr")
        row_count = min(10, rows.count())
        if verbose:
            print(f"[HSHARE] 开始解析 {row_count} 行。")

        hit_subscribe = 0
        hit_listing = 0

        for i in range(row_count):
            row = rows.nth(i)
            cells = row.locator("td")
            if cells.count() < 3:
                continue

            code, name = _extract_code_name(cells, reference_date=reference_date)
            if not code or not name:
                continue

            item = {
                "code": code,
                "name": name,
                "market": "HK",
            }

            subscribe_date = _extract_date_by_priority(
                cells,
                indexes=[3, 6],
                reference_date=reference_date,
            )
            listing_date = _extract_date_by_priority(
                cells,
                indexes=[4, 7],
                reference_date=reference_date,
            )

            if subscribe_date in date_set:
                raw_data[subscribe_date]["subscribe"].append(dict(item))
                hit_subscribe += 1

            if listing_date in date_set:
                raw_data[listing_date]["listing"].append(dict(item))
                hit_listing += 1

        if verbose:
            print(f"[HSHARE] 命中 subscribe={hit_subscribe}, listing={hit_listing}")

        browser.close()

    return normalize_fetch_output(
        raw_data,
        date_list=date_list,
        reference_date=reference_date,
    )


def main() -> None:
    hshare_map = fetch(reference_date=REFERENCE_DATE, verbose=True)
    payload = {
        "generated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "reference_date": REFERENCE_DATE,
        "date_list": get_calendar_range(REFERENCE_DATE),
        "h_share": hshare_map,
        "hk": hshare_map,
    }
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_FILE.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"[HSHARE] written: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
