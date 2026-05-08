from __future__ import annotations

import datetime as dt
import json
import re
import sys
from pathlib import Path
from typing import Any

from playwright.sync_api import Locator, Page, sync_playwright

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from utils import clean_data, get_calendar_range, get_playwright_headless, init_calendar_data, normalize_fetch_output


URL = "https://www.sse.com.cn/ipo/home/"
REFERENCE_DATE = dt.datetime.now().date().strftime("%Y-%m-%d")

DATE_UL_XPATH = "/html/body/div[5]/div/div[1]/ul"
LEFT_BTN_XPATHS = [
    "/html/body/div[5]/div/div[1]/div[1]/div[1]",
    "/html/body/div[5]/div/div[1]/div[1]",
]
RIGHT_BTN_XPATHS = [
    "/html/body/div[5]/div/div[1]/div[1]/div[2]",
    "/html/body/div[5]/div/div[1]/div[2]",
]

DRAFTING_XPATH = "/html/body/div[5]/div/div[2]/div/table/tbody/tr/td[2]/div[2]/span"
INQUIRY_XPATH = "/html/body/div[5]/div/div[2]/div/table/tbody/tr/td[3]/div[2]/div/span"
CALENDAR_CLICK_SETTLE_MS = 1500
CONTENT_SYNC_POLL_MS = 400
CONTENT_REFRESH_TIMEOUT_MS = 12000
CONTENT_STABLE_POLLS = 3


def _is_no_content_text(text: str) -> bool:
    normalized = clean_data(text).strip()
    return normalized in {"", "暂无", "无", "--"}


def _first_clickable(page: Page, xpaths: list[str]) -> Locator | None:
    for xpath in xpaths:
        locator = page.locator(f"xpath={xpath}")
        if locator.count() > 0:
            return locator.first
    return None


def _visible_date_tabs(page: Page, reference_date: str) -> dict[str, Locator]:
    result: dict[str, Locator] = {}
    ul = page.locator(f"xpath={DATE_UL_XPATH}")
    if ul.count() == 0:
        return result
    tabs = ul.first.locator("li")
    for i in range(tabs.count()):
        tab = tabs.nth(i)
        key = clean_data(tab.inner_text(), reference_date=reference_date)
        if isinstance(key, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", key):
            result[key] = tab
    return result


def _goto_target_date(page: Page, target_date: str, reference_date: str, max_steps: int = 24) -> bool:
    left_btn = _first_clickable(page, LEFT_BTN_XPATHS)
    right_btn = _first_clickable(page, RIGHT_BTN_XPATHS)

    for _ in range(max_steps):
        visible_tabs = _visible_date_tabs(page, reference_date)
        if target_date in visible_tabs:
            visible_tabs[target_date].click()
            page.wait_for_timeout(CALENDAR_CLICK_SETTLE_MS)
            return True
        if not visible_tabs:
            return False

        visible_dates = sorted(visible_tabs.keys())
        if target_date < visible_dates[0]:
            if left_btn is None:
                return False
            left_btn.click()
        elif target_date > visible_dates[-1]:
            if right_btn is None:
                return False
            right_btn.click()
        else:
            if right_btn is None:
                return False
            right_btn.click()
        page.wait_for_timeout(CALENDAR_CLICK_SETTLE_MS)

    return False


def _get_section_signature(page: Page, xpath: str, reference_date: str) -> str:
    spans = page.locator(f"xpath={xpath}")
    if spans.count() == 0:
        return "__missing__"

    parts: list[str] = []
    for i in range(spans.count()):
        span = spans.nth(i)
        cls = (span.get_attribute("class") or "").strip()
        text = clean_data(span.inner_text(), reference_date=reference_date)
        parts.append(f"{cls}::{text}")

    return f"{spans.count()}::" + "||".join(parts)


def _wait_sections_refreshed(
    page: Page,
    reference_date: str,
    previous_drafting_signature: str | None,
    previous_inquiry_signature: str | None,
    timeout_ms: int = CONTENT_REFRESH_TIMEOUT_MS,
) -> bool:
    deadline = dt.datetime.now().timestamp() + timeout_ms / 1000.0
    grace_deadline = dt.datetime.now().timestamp() + CALENDAR_CLICK_SETTLE_MS / 1000.0
    stable_pair: tuple[str, str] | None = None
    stable_count = 0
    changed = previous_drafting_signature is None or previous_inquiry_signature is None

    while dt.datetime.now().timestamp() < deadline:
        drafting_signature = _get_section_signature(
            page,
            DRAFTING_XPATH,
            reference_date=reference_date,
        )
        inquiry_signature = _get_section_signature(
            page,
            INQUIRY_XPATH,
            reference_date=reference_date,
        )
        current_pair = (drafting_signature, inquiry_signature)

        if (
            previous_drafting_signature is None
            or previous_inquiry_signature is None
            or drafting_signature != previous_drafting_signature
            or inquiry_signature != previous_inquiry_signature
        ):
            changed = True

        if current_pair == stable_pair:
            stable_count += 1
        else:
            stable_pair = current_pair
            stable_count = 1

        if stable_count >= CONTENT_STABLE_POLLS and (
            changed or dt.datetime.now().timestamp() >= grace_deadline
        ):
            return True

        page.wait_for_timeout(CONTENT_SYNC_POLL_MS)

    return False


def _extract_items_with_status(spans: Locator, reference_date: str, market: str) -> tuple[list[dict[str, str]], str]:
    span_count = spans.count()
    if span_count == 0:
        return [], "xpath_missing"

    items: list[dict[str, str]] = []
    seen: set[str] = set()
    saw_placeholder = False

    for i in range(span_count):
        span = spans.nth(i)
        cls = (span.get_attribute("class") or "").strip()
        if "tody-li-no" in cls:
            saw_placeholder = True
            continue

        links = span.locator("a")
        if links.count() > 0:
            for j in range(links.count()):
                name = clean_data(links.nth(j).inner_text(), reference_date=reference_date)
                if _is_no_content_text(str(name)):
                    saw_placeholder = True
                    continue
                if name in seen:
                    continue
                seen.add(name)
                items.append({"name": name, "market": market})
            continue

        text = clean_data(span.inner_text(), reference_date=reference_date)
        if _is_no_content_text(str(text)):
            saw_placeholder = True
            continue

        name = re.split(r"[：:]", str(text))[-1].strip()
        name = clean_data(name, reference_date=reference_date)
        if _is_no_content_text(str(name)):
            saw_placeholder = True
            continue
        if name in seen:
            continue
        seen.add(name)
        items.append({"name": name, "market": market})

    if items:
        return items, "data_found"
    if saw_placeholder:
        return [], "state_confirmed_empty"
    return [], "state_confirmed_empty"


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
        page.wait_for_selector(f"xpath={DATE_UL_XPATH}", timeout=timeout_ms)

        for target_date in date_list:
            previous_drafting_signature = _get_section_signature(
                page,
                DRAFTING_XPATH,
                reference_date=reference_date,
            )
            previous_inquiry_signature = _get_section_signature(
                page,
                INQUIRY_XPATH,
                reference_date=reference_date,
            )
            ok = _goto_target_date(page, target_date=target_date, reference_date=reference_date)
            if not ok:
                if verbose:
                    print(f"[SSE][{target_date}] 日期标签未命中，跳过。")
                continue

            content_ready = _wait_sections_refreshed(
                page,
                reference_date=reference_date,
                previous_drafting_signature=previous_drafting_signature,
                previous_inquiry_signature=previous_inquiry_signature,
                timeout_ms=CONTENT_REFRESH_TIMEOUT_MS,
            )
            if not content_ready:
                if verbose:
                    print(f"[SSE][{target_date}] drafting/inquiry regions did not finish refreshing")
                continue

            drafting_spans = page.locator(f"xpath={DRAFTING_XPATH}")
            drafting_items, drafting_status = _extract_items_with_status(
                drafting_spans,
                reference_date=reference_date,
                market="SH",
            )
            raw_data[target_date]["drafting"] = drafting_items

            inquiry_spans = page.locator(f"xpath={INQUIRY_XPATH}")
            inquiry_items, inquiry_status = _extract_items_with_status(
                inquiry_spans,
                reference_date=reference_date,
                market="SH",
            )
            raw_data[target_date]["inquiry"] = inquiry_items

            if not verbose:
                continue

            if drafting_status == "xpath_missing":
                print(f"[SSE][{target_date}][drafting] XPath未命中: {DRAFTING_XPATH}")
            elif drafting_status == "state_confirmed_empty":
                print(f"[SSE][{target_date}][drafting] 状态确认: 已命中元素，但内容为空/暂无。")
            else:
                print(f"[SSE][{target_date}][drafting] 命中有效数据: {len(drafting_items)} 条。")

            if inquiry_status == "xpath_missing":
                print(f"[SSE][{target_date}][inquiry] XPath未命中: {INQUIRY_XPATH}")
            elif inquiry_status == "state_confirmed_empty":
                print(f"[SSE][{target_date}][inquiry] 状态确认: 已命中元素，但内容为空/暂无。")
            else:
                print(f"[SSE][{target_date}][inquiry] 命中有效数据: {len(inquiry_items)} 条。")

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
