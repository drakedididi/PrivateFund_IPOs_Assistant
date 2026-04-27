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


URL_STOCK = "https://data.eastmoney.com/xg/xg"
URL_REITS = "https://data.eastmoney.com/xg/xg/?mkt=reits"

STOCK_TBODY_XPATH = '//*[@id="dataview_hs"]/div[2]/div[2]/table/tbody'
REITS_TBODY_XPATH = '//*[@id="dataview_reits"]/div[2]/div[2]/table/tbody'

REFERENCE_DATE = dt.datetime.now().date().strftime("%Y-%m-%d")


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


def _infer_market(code: str, href: str) -> str:
    m_unify = re.search(r"/r/([01])\.\d+", href)
    if m_unify:
        return "SH" if m_unify.group(1) == "1" else "SZ"

    m_quote = re.search(r"/(SH|SZ|BJ)\d+\.html", href, flags=re.IGNORECASE)
    if m_quote:
        return m_quote.group(1).upper()

    if code.startswith(("60", "68", "50", "51", "52", "53", "56", "58")):
        return "SH"
    if code.startswith(("00", "30", "12", "15", "16", "18")):
        return "SZ"
    if code.startswith(("4", "8", "9")):
        return "BJ"
    return "A"


def _extract_security_item(row: Locator, reference_date: str) -> dict[str, str] | None:
    cells = row.locator("td")
    if cells.count() < 2:
        return None

    code_cell = cells.nth(0)
    name_cell = cells.nth(1)

    code = clean_data(code_cell.inner_text(), reference_date=reference_date)
    name = clean_data(name_cell.inner_text(), reference_date=reference_date)
    if not isinstance(code, str) or not isinstance(name, str):
        return None
    code = code.strip()
    name = name.strip()
    if not code or not name:
        return None

    href = ""
    code_link = code_cell.locator("a")
    if code_link.count() > 0:
        href = code_link.first.get_attribute("href") or ""
    if not href:
        name_link = name_cell.locator("a")
        if name_link.count() > 0:
            href = name_link.first.get_attribute("href") or ""

    market = "BJ" if code.startswith("9") else _infer_market(code, href)
    return {"code": code, "name": name, "market": market}


def _new_day_bucket() -> dict[str, list[Any]]:
    return {
        "drafting": [],
        "inquiry": [],
        "subscribe": [],
        "payment": [],
        "listing": [],
    }


def _append_event(
    raw_data: dict[str, dict[str, list[Any]]],
    date_key: str,
    event_key: str,
    item: dict[str, str],
    date_set: set[str] | None = None,
) -> bool:
    if not date_key:
        return False
    if date_set is not None and date_key not in date_set:
        return False
    if date_key not in raw_data:
        raw_data[date_key] = _new_day_bucket()
    raw_data[date_key][event_key].append(dict(item))
    return True


def _parse_stock_table(
    page,
    raw_data: dict[str, dict[str, list[Any]]],
    date_set: set[str] | None,
    reference_date: str,
    limit: int = 20,
    verbose: bool = True,
) -> None:
    tbody = page.locator(f"xpath={STOCK_TBODY_XPATH}")
    if tbody.count() == 0:
        if verbose:
            print(f"[EASTMONEY][STOCK] XPath未命中: {STOCK_TBODY_XPATH}")
        return

    rows = tbody.first.locator("tr")
    row_count = min(limit, rows.count())
    if verbose:
        print(f"[EASTMONEY][STOCK] 开始解析 {row_count} 行。")

    hit_sub = 0
    hit_pay = 0
    hit_list = 0
    skip_invalid = 0
    skip_no_date = 0

    for i in range(row_count):
        row = rows.nth(i)
        item = _extract_security_item(row, reference_date=reference_date)
        if item is None:
            skip_invalid += 1
            continue

        cells = row.locator("td")
        if cells.count() < 15:
            skip_invalid += 1
            continue

        subscribe_date = _extract_date_from_td(cells.nth(11), reference_date=reference_date)
        payment_date = _extract_date_from_td(cells.nth(13), reference_date=reference_date)
        listing_date = _extract_date_from_td(cells.nth(14), reference_date=reference_date)

        appended_any = False
        if _append_event(raw_data, subscribe_date, "subscribe", item, date_set):
            hit_sub += 1
            appended_any = True
        if _append_event(raw_data, payment_date, "payment", item, date_set):
            hit_pay += 1
            appended_any = True
        if _append_event(raw_data, listing_date, "listing", item, date_set):
            hit_list += 1
            appended_any = True
        if not appended_any:
            skip_no_date += 1

        if verbose and i == 0:
            print(
                f"[EASTMONEY][STOCK][row0] {item['name']}({item['code']}) "
                f"subscribe={subscribe_date or '-'}, payment={payment_date or '-'}, listing={listing_date or '-'}"
            )

    if verbose:
        print(
            f"[EASTMONEY][STOCK] 命中 subscribe={hit_sub}, payment={hit_pay}, listing={hit_list}; "
            f"无日期/窗口外={skip_no_date}, 无效行={skip_invalid}"
        )


def _parse_reits_table(
    page,
    raw_data: dict[str, dict[str, list[Any]]],
    date_set: set[str] | None,
    reference_date: str,
    limit: int = 10,
    verbose: bool = True,
) -> None:
    tbody = page.locator(f"xpath={REITS_TBODY_XPATH}")
    if tbody.count() == 0:
        if verbose:
            print(f"[EASTMONEY][REITS] XPath未命中: {REITS_TBODY_XPATH}")
        return

    rows = tbody.first.locator("tr")
    row_count = min(limit, rows.count())
    if verbose:
        print(f"[EASTMONEY][REITS] 开始解析 {row_count} 行。")

    hit_sub = 0
    hit_list = 0
    skip_invalid = 0
    skip_no_date = 0

    for i in range(row_count):
        row = rows.nth(i)
        item = _extract_security_item(row, reference_date=reference_date)
        if item is None:
            skip_invalid += 1
            continue

        cells = row.locator("td")
        if cells.count() < 15:
            skip_invalid += 1
            continue

        subscribe_date = _extract_date_from_td(cells.nth(10), reference_date=reference_date)
        listing_date = _extract_date_from_td(cells.nth(14), reference_date=reference_date)

        appended_any = False
        if _append_event(raw_data, subscribe_date, "subscribe", item, date_set):
            hit_sub += 1
            appended_any = True
        if _append_event(raw_data, listing_date, "listing", item, date_set):
            hit_list += 1
            appended_any = True
        if not appended_any:
            skip_no_date += 1

    if verbose:
        print(
            f"[EASTMONEY][REITS] 命中 subscribe={hit_sub}, listing={hit_list}; "
            f"无日期/窗口外={skip_no_date}, 无效行={skip_invalid}"
        )


def fetch(
    reference_date: str = REFERENCE_DATE,
    timeout_ms: int = 30000,
    verbose: bool = True,
) -> dict[str, dict[str, list[Any]]]:
    base_window = get_calendar_range(reference_date)
    raw_data = init_calendar_data(base_window)
    if verbose and base_window:
        print(f"[EASTMONEY] 基础窗口: {base_window[0]} ~ {base_window[-1]}，当前模式: 不过滤日期")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        page.goto(URL_STOCK, wait_until="domcontentloaded", timeout=timeout_ms)
        page.wait_for_selector(f"xpath={STOCK_TBODY_XPATH}", timeout=timeout_ms)
        _parse_stock_table(
            page=page,
            raw_data=raw_data,
            date_set=None,
            reference_date=reference_date,
            limit=20,
            verbose=verbose,
        )

        page.goto(URL_REITS, wait_until="domcontentloaded", timeout=timeout_ms)
        page.wait_for_selector(f"xpath={REITS_TBODY_XPATH}", timeout=timeout_ms)
        _parse_reits_table(
            page=page,
            raw_data=raw_data,
            date_set=None,
            reference_date=reference_date,
            limit=10,
            verbose=verbose,
        )

        browser.close()

    output_dates = sorted(raw_data.keys())
    return normalize_fetch_output(
        raw_data,
        date_list=output_dates,
        reference_date=reference_date,
    )


def main() -> None:
    try:
        import sys

        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    data = fetch(reference_date=REFERENCE_DATE, verbose=True)
    print(json.dumps(data, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
