import argparse

from pcf_service import save_market_result


def main() -> None:
    parser = argparse.ArgumentParser(description="生成深交所 ETF 白名单结果")
    parser.add_argument("stock_code", nargs="?", help="股票代码，例如 688825.SH")
    parser.add_argument("--stock-name", default="", help="证券名称")
    parser.add_argument("--work-dir", help="项目目录外的 PCF 数据目录")
    parser.add_argument("--output", help="结果 Excel 路径")
    args = parser.parse_args()
    stock_code = args.stock_code or input("请输入股票代码：").strip()
    result = save_market_result(
        "SZ",
        stock_code,
        args.work_dir,
        args.output,
        stock_name=args.stock_name,
    )
    print(f"[SUCCESS] 匹配 {result['matches']} 个 ETF: {result['path']}")


if __name__ == "__main__":
    main()
