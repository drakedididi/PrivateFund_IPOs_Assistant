from __future__ import annotations

import copy
import datetime as dt
import json
import os
import re
from typing import Any, Iterable, Mapping, Sequence


EMPTY_DAY = {
    "drafting": [],
    "inquiry": [],
    "subscribe": [],
    "payment": [],
    "listing": [],
}

EVENT_KEYS = tuple(EMPTY_DAY.keys())
_PREFIX_RE = re.compile(r"^(?:\s*(?:N|U|W|ST)\s*)+", re.IGNORECASE)
_DATE_RE = re.compile(r"(?:(\d{4})[./-])?\s*(\d{1,2})[./-](\d{1,2})")
_HEADLESS_TRUE_VALUES = {"1", "true", "yes", "on"}
_HEADLESS_FALSE_VALUES = {"0", "false", "no", "off"}


def _to_date(value: dt.date | dt.datetime | str | None = None) -> dt.date:
    if value is None:
        return dt.datetime.now().date()
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = str(value).strip()
    return dt.datetime.strptime(text, "%Y-%m-%d").date()


def _parse_date_text(text: str, reference: dt.date) -> str:
    match = _DATE_RE.search(text)
    if not match:
        return text

    year_s, month_s, day_s = match.groups()
    month = int(month_s)
    day = int(day_s)

    if year_s:
        candidate = dt.date(int(year_s), month, day)
    else:
        candidates = []
        for y in (reference.year - 1, reference.year, reference.year + 1):
            try:
                candidates.append(dt.date(y, month, day))
            except ValueError:
                continue
        if not candidates:
            return text
        candidate = min(candidates, key=lambda d: abs((d - reference).days))

    return candidate.strftime("%Y-%m-%d")


def get_playwright_headless(default: bool = False) -> bool:
    raw = os.getenv("PLAYWRIGHT_HEADLESS")
    if raw is None:
        return default

    value = raw.strip().lower()
    if not value:
        return default
    if value in _HEADLESS_TRUE_VALUES:
        return True
    if value in _HEADLESS_FALSE_VALUES:
        return False
    return default


def get_calendar_range(reference_date: dt.date | dt.datetime | str | None = None) -> list[str]:
    ref = _to_date(reference_date)
    monday = ref - dt.timedelta(days=ref.weekday())
    current_week = [
        (monday + dt.timedelta(days=offset)).strftime("%Y-%m-%d")
        for offset in range(5)
    ]
    next_monday = monday + dt.timedelta(days=7)
    next_week = [
        (next_monday + dt.timedelta(days=offset)).strftime("%Y-%m-%d")
        for offset in range(5)
    ]
    return current_week + next_week


def init_calendar_data(date_list: Sequence[str]) -> dict[str, dict[str, list[Any]]]:
    return {date_key: copy.deepcopy(EMPTY_DAY) for date_key in date_list}


def clean_data(value: Any, reference_date: dt.date | dt.datetime | str | None = None) -> Any:
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return value

    text = str(value).replace("\u3000", " ").strip()
    text = re.sub(r"\s+", " ", text)

    ref = _to_date(reference_date)
    maybe_date = _parse_date_text(text, ref)
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", maybe_date):
        return maybe_date

    text = _PREFIX_RE.sub("", text).strip()
    return text


def in_calendar_range(date_key: str, date_list: Sequence[str]) -> bool:
    return clean_data(date_key) in set(date_list)


def normalize_fetch_output(
    raw_data: Mapping[str, Mapping[str, Iterable[Any]]] | None,
    date_list: Sequence[str] | None = None,
    reference_date: dt.date | dt.datetime | str | None = None,
) -> dict[str, dict[str, list[Any]]]:
    dates = list(date_list) if date_list else get_calendar_range(reference_date)
    normalized = init_calendar_data(dates)
    date_set = set(dates)

    if not raw_data:
        return normalized

    for raw_date, day_data in raw_data.items():
        date_key = clean_data(raw_date, reference_date=reference_date)
        if date_key not in date_set:
            continue
        if not isinstance(day_data, Mapping):
            continue

        for event_key in EVENT_KEYS:
            items = day_data.get(event_key, [])
            if items is None:
                continue
            if isinstance(items, (str, bytes)):
                items = [items]

            cleaned_items: list[Any] = []
            seen: set[str] = set()
            for item in items:
                if isinstance(item, Mapping):
                    cleaned = {}
                    for k, v in item.items():
                        if isinstance(v, (str, bytes)):
                            cleaned[k] = clean_data(v, reference_date=reference_date)
                        else:
                            cleaned[k] = v
                    signature = f"dict:{json.dumps(cleaned, ensure_ascii=False, sort_keys=True)}"
                else:
                    cleaned = clean_data(item, reference_date=reference_date)
                    if isinstance(cleaned, (int, float)):
                        cleaned = str(cleaned)
                    signature = f"scalar:{cleaned}"
                if not cleaned:
                    continue
                if signature not in seen:
                    seen.add(signature)
                    cleaned_items.append(cleaned)
            normalized[date_key][event_key] = cleaned_items

    return normalized


def filter_calendar_window(
    raw_data: Mapping[str, Mapping[str, Iterable[Any]]] | None,
    reference_date: dt.date | dt.datetime | str | None = None,
) -> dict[str, dict[str, list[Any]]]:
    date_list = get_calendar_range(reference_date)
    return normalize_fetch_output(raw_data, date_list=date_list, reference_date=reference_date)
