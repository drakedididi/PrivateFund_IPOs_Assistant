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


URL = "https://www.bse.cn/issue/issue_calendar.html"
TBODY_XPATH = '//*[@id="issue_calendar_title"]/tbody'
REFERENCE_DATE = dt.datetime.now().date().strftime("%Y-%m-%d")


def _is_no_content_text(text: str) -> bool:
    normalized = clean_data(text).strip()
    return normalized in {"", "暂无", "无", "--"}


def _is_visible(style: str | None) -> bool:
    if not style:
        return True
    normalized = style.replace(" ", "").lower()
    return "display:none" not in normalized


def _extract_from_td_with_status(
    td: Locator,
    include_class_token: str,
    reference_date: str,
    market: str,
    exclude_class_token: str | None = None,
) -> tuple[list[dict[str, str]], str]:
    entries: list[dict[str, str]] = []
    seen_names: set[str] = set()
    saw_placeholder = False
    found_block = False

    divs = td.locator("div")
    for i in range(divs.count()):
        div = divs.nth(i)
        class_attr = (div.get_attribute("class") or "").strip()
        if include_class_token not in class_attr:
            continue
        if exclude_class_token and exclude_class_token in class_attr:
            continue

        found_block = True
        if not _is_visible(div.get_attribute("style")):
            saw_placeholder = True
            continue

        context = div.locator(".span_context")
        if context.count() == 0:
            saw_placeholder = True
            continue

        context_node = context.first
        links = context_node.locator("a")
        if links.count() > 0:
            for j in range(links.count()):
                name = clean_data(links.nth(j).inner_text(), reference_date=reference_date)
                if _is_no_content_text(str(name)):
                    saw_placeholder = True
                    continue
                if name in seen_names:
                    continue
                seen_names.add(name)
                entries.append({"name": name, "market": market})
            continue

        text = clean_data(context_node.inner_text(), reference_date=reference_date)
        if _is_no_content_text(str(text)):
            saw_placeholder = True
            continue
        name = re.split(r"[：:]", str(text))[-1].strip()
        name = clean_data(name, reference_date=reference_date)
        if _is_no_content_text(str(name)):
            saw_placeholder = True
            continue
        if name in seen_names:
            continue
        seen_names.add(name)
        entries.append({"name": name, "market": market})

    if not found_block:
        return [], "xpath_missing"
    if entries:
        return entries, "data_found"
    if saw_placeholder:
        return [], "state_confirmed_empty"
    return [], "state_confirmed_empty"


def _extract_multi_from_td_with_status(
    td: Locator,
    include_class_tokens: list[str],
    reference_date: str,
    market: str,
) -> tuple[list[dict[str, str]], str]:
    entries: list[dict[str, str]] = []
    seen_names: set[str] = set()
    statuses: list[str] = []

    for class_token in include_class_tokens:
        items, status = _extract_from_td_with_status(
            td=td,
            include_class_token=class_token,
            exclude_class_token=None,
            reference_date=reference_date,
            market=market,
        )
        statuses.append(status)
        for item in items:
            name = item.get("name", "")
            if not name or name in seen_names:
                continue
            seen_names.add(name)
            entries.append(item)

    if entries:
        return entries, "data_found"
    if statuses and all(status == "xpath_missing" for status in statuses):
        return [], "xpath_missing"
    return [], "state_confirmed_empty"


def _reserved_visibility(td: Locator) -> dict[str, bool]:
    targets = {
        "subscribe": ["title_purchase_"],
        "payment": ["title_issue_result_"],
        "listing": ["title_enter_premium_"],
    }
    result: dict[str, bool] = {}

    for key, class_keys in targets.items():
        visible = False
        for class_key in class_keys:
            blocks = td.locator(f"div[class*='{class_key}']")
            for i in range(blocks.count()):
                if _is_visible(blocks.nth(i).get_attribute("style")):
                    visible = True
                    break
            if visible:
                break
        result[key] = visible

    return result


def fetch(
    reference_date: str = REFERENCE_DATE,
    timeout_ms: int = 30000,
    verbose: bool = True,
) -> dict[str, dict[str, list[Any]]]:
    date_list = get_calendar_range(reference_date)
    raw_data = init_calendar_data(date_list)

    row_date_map = {
        2: date_list[0:5],
        6: date_list[5:10],
    }

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=get_playwright_headless())
        page = browser.new_page()
        page.goto(URL, wait_until="networkidle", timeout=timeout_ms)
        page.wait_for_selector(f"xpath={TBODY_XPATH}", timeout=timeout_ms)
        page.wait_for_function(
            """
            ([startDate, endDate]) => {
                const tbody = document.querySelector("#issue_calendar_title tbody");
                if (!tbody) {
                    return false;
                }
                const text = tbody.innerText || "";
                return text.includes(startDate) && text.includes(endDate);
            }
            """,
            arg=[date_list[0], date_list[-1]],
            timeout=timeout_ms,
        )

        tbody = page.locator(f"xpath={TBODY_XPATH}")
        rows = tbody.locator("tr")
        row_count = rows.count()

        for row_index, date_keys in row_date_map.items():
            if row_index >= row_count:
                continue

            row = rows.nth(row_index)
            tds = row.locator("td")
            slot_count = min(5, tds.count(), len(date_keys))

            for col in range(slot_count):
                date_key = clean_data(date_keys[col], reference_date=reference_date)
                if date_key not in raw_data:
                    continue

                td = tds.nth(col)
                drafting_items, drafting_status = _extract_from_td_with_status(
                    td=td,
                    include_class_token="title_enquiry_notice_",
                    exclude_class_token=None,
                    reference_date=reference_date,
                    market="BJ",
                )
                raw_data[date_key]["drafting"] = drafting_items

                inquiry_items, inquiry_status = _extract_from_td_with_status(
                    td=td,
                    include_class_token="title_enquiry_",
                    exclude_class_token="title_enquiry_notice_",
                    reference_date=reference_date,
                    market="BJ",
                )
                raw_data[date_key]["inquiry"] = inquiry_items

                subscribe_items, subscribe_status = _extract_from_td_with_status(
                    td=td,
                    include_class_token="title_purchase_",
                    exclude_class_token=None,
                    reference_date=reference_date,
                    market="BJ",
                )
                raw_data[date_key]["subscribe"] = subscribe_items

                payment_items, payment_status = _extract_multi_from_td_with_status(
                    td=td,
                    include_class_tokens=["title_issue_result_"],
                    reference_date=reference_date,
                    market="BJ",
                )
                raw_data[date_key]["payment"] = payment_items

                listing_items, listing_status = _extract_from_td_with_status(
                    td=td,
                    include_class_token="title_enter_premium_",
                    exclude_class_token=None,
                    reference_date=reference_date,
                    market="BJ",
                )
                raw_data[date_key]["listing"] = listing_items

                if verbose:
                    if drafting_status == "xpath_missing":
                        print(f"[BSE][{date_key}][drafting] XPath未命中: div[class*='title_enquiry_notice_']")
                    elif drafting_status == "state_confirmed_empty":
                        print(f"[BSE][{date_key}][drafting] 状态确认: 已命中元素，但内容为空/暂无/隐藏。")
                    else:
                        print(f"[BSE][{date_key}][drafting] 命中有效数据: {len(drafting_items)} 条。")

                    if inquiry_status == "xpath_missing":
                        print(f"[BSE][{date_key}][inquiry] XPath未命中: div[class*='title_enquiry_']")
                    elif inquiry_status == "state_confirmed_empty":
                        print(f"[BSE][{date_key}][inquiry] 状态确认: 已命中元素，但内容为空/无/隐藏。")
                    else:
                        print(f"[BSE][{date_key}][inquiry] 命中有效数据: {len(inquiry_items)} 条。")

                    if subscribe_status == "xpath_missing":
                        print(f"[BSE][{date_key}][subscribe] XPath未命中: div[class*='title_purchase_']")
                    elif subscribe_status == "state_confirmed_empty":
                        print(f"[BSE][{date_key}][subscribe] 状态确认: 已命中元素，但内容为空/无/隐藏。")
                    else:
                        print(f"[BSE][{date_key}][subscribe] 命中有效数据: {len(subscribe_items)} 条。")

                    if payment_status == "xpath_missing":
                        print(f"[BSE][{date_key}][payment] XPath未命中: div[class*='title_issue_result_']")
                    elif payment_status == "state_confirmed_empty":
                        print(f"[BSE][{date_key}][payment] 状态确认: 已命中元素，但内容为空/无/隐藏。")
                    else:
                        print(f"[BSE][{date_key}][payment] 命中有效数据: {len(payment_items)} 条。")

                    if listing_status == "xpath_missing":
                        print(f"[BSE][{date_key}][listing] XPath未命中: div[class*='title_enter_premium_']")
                    elif listing_status == "state_confirmed_empty":
                        print(f"[BSE][{date_key}][listing] 状态确认: 已命中元素，但内容为空/无/隐藏。")
                    else:
                        print(f"[BSE][{date_key}][listing] 命中有效数据: {len(listing_items)} 条。")

                _ = _reserved_visibility(td)

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
