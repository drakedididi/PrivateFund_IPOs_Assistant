from datetime import datetime

print(f"开始爬取，当前时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

from merge.scramer_merge import main


if __name__ == "__main__":
    main()
