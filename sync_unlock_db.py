from datetime import datetime

print(f"开始同步unlock_db，当前时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

from merge.scramer_merge import sync_unlock_db_main


if __name__ == "__main__":
    sync_unlock_db_main()
