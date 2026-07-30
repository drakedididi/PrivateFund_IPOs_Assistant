from __future__ import annotations

import argparse
import json
import os
import socket
import sys
from pathlib import Path
from urllib.request import ProxyHandler, build_opener


def find_data_dir(repo_root: Path, configured: str) -> Path:
    if configured:
        return Path(configured).expanduser().resolve()
    for base in repo_root.parents:
        candidate = base / "pcf_runtime"
        if (candidate / "SH").exists() or (candidate / "SZ").exists():
            return candidate.resolve()
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share"))
    return (base / "PrivateFundIPOsAssistant" / "pcf").resolve()


def port_is_open(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as connection:
        connection.settimeout(0.3)
        return connection.connect_ex(("127.0.0.1", port)) == 0


def find_available_port(preferred_port: int) -> int:
    port = preferred_port
    while port_is_open(port):
        port += 1
        if port > 65535:
            raise RuntimeError("没有找到可用的本地端口")
    return port


def current_pcf_service_url(port: int) -> str:
    url = f"http://127.0.0.1:{port}"
    try:
        opener = build_opener(ProxyHandler({}))
        with opener.open(f"{url}/api/pcf/status", timeout=10) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except (OSError, ValueError, json.JSONDecodeError):
        return ""
    if payload.get("work_dir") and payload.get("log_path"):
        return url
    return ""


def main() -> None:
    parser = argparse.ArgumentParser(description="启动本地 PCF 白名单网页和 API")
    parser.add_argument("--port", type=int, default=5000)
    parser.add_argument("--data-dir", default=os.environ.get("PCF_DATA_DIR", ""))
    parser.add_argument("--token", default=os.environ.get("APP_SECRET_TOKEN", "local-test"))
    args = parser.parse_args()

    repo_root = Path(__file__).resolve().parent
    data_dir = find_data_dir(repo_root, args.data_dir)
    data_dir.mkdir(parents=True, exist_ok=True)
    for candidate_port in range(args.port, min(args.port + 6, 65536)):
        running_url = current_pcf_service_url(candidate_port)
        if running_url:
            print(f"当前版本已经运行，请直接打开：{running_url}/frontend/trade_tools.html")
            return

    selected_port = find_available_port(args.port)
    page_url = f"http://127.0.0.1:{selected_port}/frontend/trade_tools.html"
    if selected_port != args.port:
        print(f"端口 {args.port} 已被占用，自动改用端口 {selected_port}。")

    os.environ["PORT"] = str(selected_port)
    os.environ["PCF_DATA_DIR"] = str(data_dir)
    os.environ["APP_SECRET_TOKEN"] = args.token

    from app import app

    print(f"PCF 数据目录：{data_dir}")
    print(f"本地访问密码：{args.token}")
    print(f"打开网页：{page_url}")
    print("按 Ctrl+C 停止本地服务。")
    app.run(
        host="127.0.0.1",
        port=selected_port,
        threaded=True,
        use_reloader=False,
    )


if __name__ == "__main__":
    main()
