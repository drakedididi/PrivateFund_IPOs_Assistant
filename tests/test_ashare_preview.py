from __future__ import annotations

import unittest

from crawlers.eastmoney_stock_detail import StockDetailFetchError
from merge.Ashare_merge import (
    _build_new_stock_preview,
    _collect_preview_candidates,
    _enrich_calendar_names,
)


def _day(**events):
    return {
        "drafting": events.get("drafting", []),
        "inquiry": events.get("inquiry", []),
        "subscribe": events.get("subscribe", []),
        "payment": events.get("payment", []),
        "listing": events.get("listing", []),
        "grey_market": [],
    }


class AsharePreviewTests(unittest.TestCase):
    def setUp(self) -> None:
        self.source = {
            "2026-07-10": _day(
                inquiry=[{"name": "完整五字名称", "market": "SZ"}],
                listing=[{"name": "已上市公司", "code": "600001", "market": "SH"}],
            ),
            "2026-07-20": _day(
                subscribe=[
                    {"name": "完整五字名称", "code": "301500", "market": "SZ"},
                    {"name": "沪市公司", "code": "603468", "market": "SH"},
                    {"name": "REIT", "code": "180101", "market": "SZ"},
                    {"name": "北交公司", "code": "920072", "market": "BJ"},
                ]
            ),
            "2026-07-24": _day(
                listing=[{"name": "完整五字名称", "code": "301500", "market": "SZ"}]
            ),
        }

    def test_candidates_match_unlisted_hs_board_scope(self) -> None:
        candidates = _collect_preview_candidates(self.source, "2026-07-17")

        self.assertEqual([item["code"] for item in candidates], ["301500", "603468"])
        self.assertEqual(candidates[0]["listing"], "2026-07-24")

    def test_build_preserves_nulls_when_detail_is_unavailable(self) -> None:
        def fake_fetcher(code, **kwargs):
            if code == "603468":
                raise StockDetailFetchError("temporary failure")
            return {
                "code": code,
                "name": "完整五字名称股份",
                "market": "SZ",
                "industry": "专用设备制造业",
                "offline_placing_num": 22_150_000,
                "issue_num": 11_000_000,
            }

        previews = _build_new_stock_preview(self.source, "2026-07-17", fetcher=fake_fetcher)

        self.assertEqual(previews[0]["name"], "完整五字名称股份")
        self.assertEqual(previews[0]["offline_placing_num"], 22_150_000)
        self.assertIsNone(previews[1]["industry"])
        self.assertIsNone(previews[1]["issue_num"])

    def test_enriches_calendar_with_full_detail_name(self) -> None:
        _enrich_calendar_names(
            self.source,
            [{"code": "301500", "name": "完整五字名称股份有限公司"}],
        )

        item = self.source["2026-07-20"]["subscribe"][0]
        self.assertEqual(item["name"], "完整五字名称股份有限公司")
        early_item = self.source["2026-07-10"]["inquiry"][0]
        self.assertEqual(early_item["name"], "完整五字名称股份有限公司")
        self.assertEqual(early_item["code"], "301500")


if __name__ == "__main__":
    unittest.main()
