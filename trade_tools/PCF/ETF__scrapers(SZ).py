import argparse

from pcf_service import scrape_links


def main() -> None:
    parser = argparse.ArgumentParser(description="抓取深交所 ETF PCF 链接")
    parser.add_argument("--work-dir", help="项目目录外的 PCF 数据目录")
    parser.add_argument("--headed", action="store_true", help="显示浏览器窗口")
    args = parser.parse_args()
    result = scrape_links("SZ", args.work_dir, headless=not args.headed)
    print(f"[SUCCESS] 已保存 {result['links']} 条链接: {result['path']}")


if __name__ == "__main__":
    main()
