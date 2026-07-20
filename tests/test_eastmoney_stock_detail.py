from __future__ import annotations

import unittest

from crawlers.eastmoney_stock_detail import (
    _extract_info_payload,
    _to_share_count,
)


DETAIL_HTML = """
<!doctype html>
<html>
  <head>
    <script>var unrelated = {"value": 1};</script>
    <script>
      var info = {
        "SECURITY_CODE": "301583",
        "SECURITY_NAME": "华润新能源科技股份有限公司",
        "TRADE_MARKET": "深圳证券交易所",
        "ISSUE_NUM": 4636.8423,
        "OFFLINE_PLACING_NUM": 16228897
      };
    </script>
  </head>
</html>
"""


class EastmoneyStockDetailTests(unittest.TestCase):
    def test_extracts_embedded_info_as_json(self) -> None:
        payload = _extract_info_payload(DETAIL_HTML)

        self.assertEqual(payload["SECURITY_CODE"], "301583")
        self.assertEqual(payload["SECURITY_NAME"], "华润新能源科技股份有限公司")

    def test_converts_eastmoney_issue_units_to_shares(self) -> None:
        self.assertEqual(_to_share_count(4636.8423, scale=10_000), 46_368_423)
        self.assertEqual(_to_share_count("16,228,897".replace(",", "")), 16_228_897)
        self.assertIsNone(_to_share_count(None))
        self.assertIsNone(_to_share_count("-"))


if __name__ == "__main__":
    unittest.main()
