import argparse

from pcf_service import download_pcf_files


def main() -> None:
    parser = argparse.ArgumentParser(description="下载上交所 ETF PCF 文本")
    parser.add_argument("--work-dir", help="项目目录外的 PCF 数据目录")
    parser.add_argument("--force", action="store_true", help="覆盖已有文件")
    parser.add_argument("--workers", type=int, default=6, help="并发下载数")
    parser.add_argument("--limit", type=int, help="仅下载前 N 个文件，用于抽样测试")
    args = parser.parse_args()
    result = download_pcf_files(
        "SH",
        args.work_dir,
        force=args.force,
        workers=args.workers,
        limit=args.limit,
    )
    print(
        f"[SUCCESS] 新增 {result['downloaded']}，跳过 {result['skipped']}，"
        f"失败 {result['failed']}: {result['path']}"
    )


if __name__ == "__main__":
    main()
