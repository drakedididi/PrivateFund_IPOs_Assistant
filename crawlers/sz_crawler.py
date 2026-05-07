from __future__ import annotations

import datetime as dt
import json
import re
import sys
import time
from pathlib import Path
from typing import Any

from playwright.sync_api import Locator, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from utils import clean_data, get_calendar_range, get_playwright_headless, init_calendar_data, normalize_fetch_output


URL = "https://www.cninfo.com.cn/eipo/index.html#/"
REFERENCE_DATE = dt.datetime.now().date().strftime("%Y-%m-%d")

IFRAME_XPATH = '//iframe[contains(@src, "about:blank")]'
CALENDAR_TBODY_XPATH = "/html/body/div/div[3]/table/tbody"
VIEW_DATE_XPATH = '//*[@id="app"]/div/div/div/div[2]/div[1]/div/div[1]/span'
DRAFTING_ROWS_XPATH = "//div[text()='招股公告']/..//table//tr"
INQUIRY_ROWS_XPATH = "//*[@id='app']/div/div/div/div[2]/div[1]/div/div[2]/div[2]/div[2]/table/tr"
CALENDAR_CLICK_SETTLE_MS = 1500
CALENDAR_SYNC_POLL_MS = 400


def _onclick_token(date_key: str) -> str:
    d = dt.datetime.strptime(date_key, "%Y-%m-%d").date()
    return f"{d.year},{d.month},{d.day}"


def _is_no_content_text(text: str) -> bool:
    normalized = clean_data(text).strip()
    return normalized in {"", "暂无", "无", "--"}


def _extract_date_from_view(text: str, target_date: str) -> str:
    src = (text or "").strip()
    m_full = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", src)
    if m_full:
        y, m, d = map(int, m_full.groups())
        return f"{y:04d}-{m:02d}-{d:02d}"
    m_md = re.search(r"(\d{1,2})\D+(\d{1,2})", src)
    if m_md:
        year = int(target_date[:4])
        m, d = map(int, m_md.groups())
        return f"{year:04d}-{m:02d}-{d:02d}"
    return ""


def _parse_onclick_date(onclick_value: str | None) -> str:
    if not onclick_value:
        return ""
    m = re.search(r"(\d{4})\s*,\s*0?(\d{1,2})\s*,\s*0?(\d{1,2})", onclick_value)
    if not m:
        return ""
    y, mth, day = map(int, m.groups())
    return f"{y:04d}-{mth:02d}-{day:02d}"


def _wait_iframe_calendar_ready(page: Page, timeout_ms: int = 10000) -> bool:
    deadline = time.monotonic() + timeout_ms / 1000.0
    while time.monotonic() < deadline:
        frames = page.locator(f"xpath={IFRAME_XPATH}")
        for i in range(frames.count()):
            frame_handle = frames.nth(i).element_handle()
            if frame_handle is None:
                continue
            frame = frame_handle.content_frame()
            if frame is None:
                continue

            tbody = frame.locator(f"xpath={CALENDAR_TBODY_XPATH}")
            if tbody.count() == 0:
                continue
            if tbody.first.locator("xpath=.//td[@onclick]").count() > 0:
                return True
        page.wait_for_timeout(CALENDAR_SYNC_POLL_MS)
    return False


def _find_calendar_cell(page: Page, date_key: str) -> tuple[Locator | None, str]:
    frames = page.locator(f"xpath={IFRAME_XPATH}")
    frame_count = frames.count()
    day_num = str(int(date_key[-2:]))
    day_pad = date_key[-2:]

    for i in range(frame_count):
        frame_handle = frames.nth(i).element_handle()
        if frame_handle is None:
            continue
        frame = frame_handle.content_frame()
        if frame is None:
            continue

        tbody = frame.locator(f"xpath={CALENDAR_TBODY_XPATH}")
        if tbody.count() == 0:
            continue

        cells = tbody.first.locator("xpath=.//td[@onclick]")
        for j in range(cells.count()):
            cell = cells.nth(j)
            onclick_value = cell.get_attribute("onclick")
            if _parse_onclick_date(onclick_value) == date_key:
                return cell, "clickable_found"

        non_clickable = tbody.first.locator(
            f"xpath=.//td[not(@onclick) and "
            f"(normalize-space(text())='{day_num}' or normalize-space(text())='{day_pad}')]"
        )
        if non_clickable.count() > 0:
            return None, "date_found_not_clickable"

    return None, "date_not_found"


def _click_day_in_iframe(page: Page, date_key: str, timeout_ms: int = 8000) -> tuple[bool, str]:
    cell, status = _find_calendar_cell(page, date_key)
    if cell is None:
        return False, status
    try:
        cell.click(timeout=timeout_ms)
        return True, "clicked"
    except PlaywrightTimeoutError:
        return False, "click_timeout"


def _wait_view_synced(page: Page, target_date: str, timeout_ms: int = 8000) -> bool:
    view = page.locator(f"xpath={VIEW_DATE_XPATH}")
    deadline = time.monotonic() + timeout_ms / 1000.0
    while time.monotonic() < deadline:
        if view.count() == 0:
            page.wait_for_timeout(CALENDAR_SYNC_POLL_MS)
            continue
        raw = view.first.inner_text()
        parsed = _extract_date_from_view(raw, target_date=target_date)
        if parsed == target_date:
            return True
        page.wait_for_timeout(CALENDAR_SYNC_POLL_MS)
    return False


def _extract_items_from_rows(
    page: Page,
    rows_xpath: str,
    reference_date: str,
    market: str,
) -> tuple[list[dict[str, str]], str, int]:
    rows = page.locator(f"xpath={rows_xpath}")
    row_count = rows.count()
    if row_count == 0:
        return [], "xpath_missing", 0

    items: list[dict[str, str]] = []
    seen: set[str] = set()
    nbsp_count = 0
    saw_cell = False

    for i in range(row_count):
        row = rows.nth(i)
        cells = row.locator("td[name='td_item']")
        for j in range(cells.count()):
            saw_cell = True
            cell = cells.nth(j)

            links = cell.locator("a")
            if links.count() > 0:
                for k in range(links.count()):
                    name = clean_data(links.nth(k).inner_text(), reference_date=reference_date)
                    if _is_no_content_text(str(name)):
                        continue
                    if name in seen:
                        continue
                    seen.add(name)
                    items.append({"name": name, "market": market})
                continue

            text_raw = (cell.inner_text() or "").replace("\xa0", " ").strip()
            text = clean_data(text_raw, reference_date=reference_date)
            if _is_no_content_text(str(text)):
                nbsp_count += 1
                continue

            name = re.split(r"[：:]", str(text))[-1].strip()
            name = clean_data(name, reference_date=reference_date)
            if _is_no_content_text(str(name)):
                nbsp_count += 1
                continue
            if name in seen:
                continue
            seen.add(name)
            items.append({"name": name, "market": market})

    if items:
        return items, "data_found", nbsp_count
    if saw_cell:
        return [], "state_confirmed_empty", nbsp_count
    return [], "xpath_missing", nbsp_count


def _capture_and_log_for_date(
    page: Page,
    target_date: str,
    raw_data: dict[str, dict[str, list[Any]]],
    reference_date: str,
    verbose: bool,
) -> None:
    drafting_items, drafting_status, drafting_nbsp = _extract_items_from_rows(
        page=page,
        rows_xpath=DRAFTING_ROWS_XPATH,
        reference_date=reference_date,
        market="SZ",
    )
    raw_data[target_date]["drafting"] = drafting_items

    inquiry_items, inquiry_status, inquiry_nbsp = _extract_items_from_rows(
        page=page,
        rows_xpath=INQUIRY_ROWS_XPATH,
        reference_date=reference_date,
        market="SZ",
    )
    raw_data[target_date]["inquiry"] = inquiry_items

    if not verbose:
        return

    if drafting_status == "xpath_missing":
        print(f"[SZSE][{target_date}][drafting] XPath未命中元素: {DRAFTING_ROWS_XPATH}")
    elif drafting_status == "state_confirmed_empty":
        print(
            f"[SZSE][{target_date}][drafting] 状态确认: XPath已命中<td name='td_item'>，"
            f"但内容为空/暂无。nbsp空位={drafting_nbsp}"
        )
    else:
        print(f"[SZSE][{target_date}][drafting] 命中有效数据: {len(drafting_items)} 条。nbsp空位={drafting_nbsp}")

    if inquiry_status == "xpath_missing":
        print(f"[SZSE][{target_date}][inquiry] XPath未命中元素: {INQUIRY_ROWS_XPATH}")
    elif inquiry_status == "state_confirmed_empty":
        print(
            f"[SZSE][{target_date}][inquiry] 状态确认: XPath已命中<td name='td_item'>，"
            f"但内容为空/暂无。nbsp空位={inquiry_nbsp}"
        )
    else:
        print(f"[SZSE][{target_date}][inquiry] 命中有效数据: {len(inquiry_items)} 条。nbsp空位={inquiry_nbsp}")


def fetch(
    reference_date: str = REFERENCE_DATE,
    timeout_ms: int = 30000,
    verbose: bool = True,
) -> dict[str, dict[str, list[Any]]]:
    date_list = get_calendar_range(reference_date)
    raw_data = init_calendar_data(date_list)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=get_playwright_headless())
        page = browser.new_page()
        page.goto(URL, wait_until="domcontentloaded", timeout=timeout_ms)
        page.wait_for_selector(f"xpath={IFRAME_XPATH}", timeout=timeout_ms)
        page.wait_for_selector(f"xpath={VIEW_DATE_XPATH}", timeout=timeout_ms)
        ready = _wait_iframe_calendar_ready(page, timeout_ms=10000)
        if not ready:
            if verbose:
                print(f"[SZSE] iframe日历未就绪: {CALENDAR_TBODY_XPATH} 或 td[@onclick] 未命中。")
            browser.close()
            return normalize_fetch_output(raw_data, date_list=date_list, reference_date=reference_date)

        first_date = date_list[0]
        first_clicked, first_click_status = _click_day_in_iframe(page, first_date, timeout_ms=8000)
        if not first_clicked:
            if verbose:
                if first_click_status == "date_found_not_clickable":
                    print(f"[SZSE][{first_date}] 日历已找到该日期，但单元格不可点击（无 onclick）。")
                elif first_click_status == "date_not_found":
                    print(f"[SZSE][{first_date}] 未在iframe日历视图中找到该日期单元格。")
                else:
                    print(f"[SZSE][{first_date}] 点击日期失败，状态={first_click_status}。")
            browser.close()
            return normalize_fetch_output(raw_data, date_list=date_list, reference_date=reference_date)

        page.wait_for_timeout(CALENDAR_CLICK_SETTLE_MS)
        first_synced = _wait_view_synced(page, target_date=first_date, timeout_ms=12000)
        if not first_synced:
            if verbose:
                print(f"[SZSE][{first_date}] 视图未同步到本周一，终止本轮抓取。")
            browser.close()
            return normalize_fetch_output(raw_data, date_list=date_list, reference_date=reference_date)

        _capture_and_log_for_date(
            page=page,
            target_date=first_date,
            raw_data=raw_data,
            reference_date=reference_date,
            verbose=verbose,
        )

        for target_date in date_list[1:]:
            clicked, click_status = _click_day_in_iframe(page, target_date, timeout_ms=8000)
            if not clicked:
                if verbose:
                    if click_status == "date_found_not_clickable":
                        print(f"[SZSE][{target_date}] 日历已找到该日期，但单元格不可点击（无 onclick）。")
                    elif click_status == "date_not_found":
                        print(f"[SZSE][{target_date}] 未在iframe日历视图中找到该日期单元格。")
                    else:
                        print(f"[SZSE][{target_date}] 点击日期失败，状态={click_status}。")
                continue

            page.wait_for_timeout(CALENDAR_CLICK_SETTLE_MS)
            synced = _wait_view_synced(page, target_date=target_date, timeout_ms=12000)
            if not synced:
                if verbose:
                    print(f"[SZSE][{target_date}] 视图未同步到目标日期，跳过抓取。")
                continue

            _capture_and_log_for_date(
                page=page,
                target_date=target_date,
                raw_data=raw_data,
                reference_date=reference_date,
                verbose=verbose,
            )

        browser.close()

    return normalize_fetch_output(
        raw_data,
        date_list=date_list,
        reference_date=reference_date,
    )


def main() -> None:
    data = fetch(reference_date=REFERENCE_DATE, verbose=True)
    print(json.dumps(data, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
