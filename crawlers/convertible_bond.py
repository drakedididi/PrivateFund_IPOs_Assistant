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

from utils import clean_data, get_calendar_range, get_playwright_headless, init_calendar_data, normalize_fetch_output


URL = "https://data.eastmoney.com/xg/xg/?mkt=kzz"
TBODY_XPATH = '//*[@id="dataview_kzz"]/div[2]/div[2]/table/tbody'

REFERENCE_DATE = dt.datetime.now().date().strftime("%Y-%m-%d")
DATA_DIR = ROOT_DIR / "data"
OUTPUT_FILE = DATA_DIR / "bondcalendar_data.json"


def _is_valid_date_key(value: str) -> bool:
    return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", value))


def _extract_date_from_td(td: Locator, reference_date: str) -> str:
    span = td.locator("span[title]")
    if span.count() > 0:
        title = span.first.get_attribute("title") or ""
        parsed = clean_data(title, reference_date=reference_date)
        if isinstance(parsed, str) and _is_valid_date_key(parsed):
            return parsed

    parsed_text = clean_data(td.inner_text(), reference_date=reference_date)
    if isinstance(parsed_text, str) and _is_valid_date_key(parsed_text):
        return parsed_text
    return ""


def _infer_market(stock_code: str) -> str:
    code = stock_code.strip()
    if code.startswith("9"):
        return "BJ"
    if code.startswith(("60", "68", "51", "52", "53", "56", "58")):
        return "SH"
    if code.startswith(("00", "30", "12", "15", "16", "18")):
        return "SZ"
    if code.startswith(("4", "8")):
        return "BJ"
    return "A"


def _extract_title_name(td: Locator, reference_date: str) -> str:
    span = td.locator("span[title]")
    if span.count() > 0:
        title = clean_data(span.first.get_attribute("title") or "", reference_date=reference_date)
        if isinstance(title, str) and title.strip():
            return title.strip()
    text = clean_data(td.inner_text(), reference_date=reference_date)
    return str(text).strip() if isinstance(text, str) else ""


def _build_item(
    bond_code: str,
    bond_name: str,
    stock_code: str,
    stock_name: str,
) -> dict[str, str]:
    display_name = f"{bond_code}（{bond_name}）+{stock_code}（{stock_name}）"
    return {
        "code": bond_code,
        "name": display_name,
        "market": _infer_market(stock_code),
    }


def fetch(
    reference_date: str = REFERENCE_DATE,
    timeout_ms: int = 30000,
    verbose: bool = True,
) -> dict[str, dict[str, list[Any]]]:
    date_list = get_calendar_range(reference_date)
    date_set = set(date_list)
    raw_data = init_calendar_data(date_list)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=get_playwright_headless())
        page = browser.new_page()
        page.goto(URL, wait_until="domcontentloaded", timeout=timeout_ms)
        page.wait_for_selector(f"xpath={TBODY_XPATH}", timeout=timeout_ms)

        tbody = page.locator(f"xpath={TBODY_XPATH}")
        rows = tbody.first.locator("tr")
        row_count = min(10, rows.count())
        if verbose:
            print(f"[BOND] 开始解析 {row_count} 行。")

        hit_subscribe = 0
        hit_listing = 0

        for i in range(row_count):
            row = rows.nth(i)
            cells = row.locator("td")
            if cells.count() < 20:
                continue

            bond_code = str(clean_data(cells.nth(0).inner_text(), reference_date=reference_date)).strip()
            bond_name = str(clean_data(cells.nth(1).inner_text(), reference_date=reference_date)).strip()
            stock_code = str(clean_data(cells.nth(6).inner_text(), reference_date=reference_date)).strip()
            stock_name = _extract_title_name(cells.nth(7), reference_date=reference_date)

            if not bond_code or not bond_name or not stock_code or not stock_name:
                continue

            item = _build_item(
                bond_code=bond_code,
                bond_name=bond_name,
                stock_code=stock_code,
                stock_name=stock_name,
            )

            subscribe_date = _extract_date_from_td(cells.nth(3), reference_date=reference_date)
            listing_date = _extract_date_from_td(cells.nth(19), reference_date=reference_date)

            if subscribe_date in date_set:
                raw_data[subscribe_date]["subscribe"].append(dict(item))
                hit_subscribe += 1
            if listing_date in date_set:
                raw_data[listing_date]["listing"].append(dict(item))
                hit_listing += 1

        if verbose:
            print(f"[BOND] 命中 subscribe={hit_subscribe}, listing={hit_listing}")

        browser.close()

    return normalize_fetch_output(
        raw_data,
        date_list=date_list,
        reference_date=reference_date,
    )


def main() -> None:
    bond_map = fetch(reference_date=REFERENCE_DATE, verbose=True)
    payload = {
        "generated_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "reference_date": REFERENCE_DATE,
        "date_list": get_calendar_range(REFERENCE_DATE),
        "bond": bond_map,
        "convertible_bond": bond_map,
    }
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_FILE.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"[BOND] written: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
