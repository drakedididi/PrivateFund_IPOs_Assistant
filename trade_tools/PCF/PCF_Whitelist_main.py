import argparse

from pathlib import Path

from pcf_service import run_pipeline


def main() -> None:
    parser = argparse.ArgumentParser(description="抓取沪深 ETF PCF 并生成白名单结果")
    parser.add_argument("stock_code", help="股票代码，例如 688825.SH")
    parser.add_argument("--stock-name", default="", help="证券名称")
    parser.add_argument("--work-dir", help="项目目录外的 PCF 数据目录")
    parser.add_argument("--market", action="append", choices=("SH", "SZ"), help="可重复指定市场")
    parser.add_argument("--refresh", action="store_true", help="重新抓取并覆盖已有 PCF")
    parser.add_argument("--output", help="结果 Excel 路径")
    args = parser.parse_args()
    result = run_pipeline(
        args.stock_code,
        args.stock_name,
        args.work_dir,
        markets=args.market or ("SH", "SZ"),
        refresh=args.refresh,
    )
    output = args.output or str(
        Path(result["work_dir"]) / f"pcf_whitelist_{result['stock_code'].replace('.', '_')}.xlsx"
    )
    result["workbook"].save(output)
    print(f"[SUCCESS] 匹配 {result['matches']} 个 ETF: {output}")
    if result["errors"]:
        print("[WARN] " + "；".join(result["errors"]))


if __name__ == "__main__":
    main()
