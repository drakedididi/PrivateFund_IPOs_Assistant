from __future__ import annotations

import argparse
import importlib.util
import io
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any

LEGACY_SCRIPT = "Invoice Organization.py"
DEFAULT_SUFFIX = "_classified.zip"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _load_legacy_module() -> Any:
    script_path = os.path.join(BASE_DIR, LEGACY_SCRIPT)
    spec = importlib.util.spec_from_file_location("invoice_legacy", script_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载脚本: {script_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _recover_mojibake(text: str) -> str:
    if not isinstance(text, str) or not text:
        return ""
    candidates = [text]
    for enc in ("gbk", "cp936"):
        try:
            repaired = text.encode(enc).decode("utf-8")
            if repaired:
                candidates.append(repaired)
        except Exception:
            pass
    return max(candidates, key=lambda s: sum(1 for ch in s if "\u4e00" <= ch <= "\u9fff"))


def _normalize_text(text: str) -> str:
    text = text.strip()
    text = re.sub(r"[\s\-—－（）()【】\[\]·,，。:：/\\]", "", text)
    return text


def _load_product_map() -> dict[str, str]:
    module = _load_legacy_module()
    raw_map = getattr(module, "PRODUCT_TO_CUSTODIAN", {})
    if not isinstance(raw_map, dict):
        raise RuntimeError("原脚本缺少 PRODUCT_TO_CUSTODIAN 映射")

    repaired: dict[str, str] = {}
    for product, custodian in raw_map.items():
        p = _recover_mojibake(str(product))
        c = _recover_mojibake(str(custodian))
        if p and c:
            repaired[p] = c
    return repaired


def _decode_zip_name(name: str) -> str:
    try:
        repaired = name.encode("cp437").decode("gbk")
    except Exception:
        return name
    old_cjk = sum(1 for ch in name if "\u4e00" <= ch <= "\u9fff")
    new_cjk = sum(1 for ch in repaired if "\u4e00" <= ch <= "\u9fff")
    return repaired if new_cjk > old_cjk else name


def _safe_target_path(base_dir: Path, member_name: str) -> Path:
    member_path = Path(member_name)
    target = (base_dir / member_path).resolve()
    base = base_dir.resolve()
    if base != target and base not in target.parents:
        raise ValueError(f"非法路径: {member_name}")
    return target


def _extract_zip(zip_bytes: bytes, extract_dir: Path) -> int:
    count = 0
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for info in zf.infolist():
            fixed_name = _decode_zip_name(info.filename)
            if not fixed_name or fixed_name.endswith("/"):
                continue
            target_path = _safe_target_path(extract_dir, fixed_name)
            target_path.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info, "r") as src, open(target_path, "wb") as dst:
                shutil.copyfileobj(src, dst)
            count += 1
    return count


def _collect_pdfs(root_dir: Path) -> list[Path]:
    return [p for p in root_dir.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf"]


def _build_matcher(product_map: dict[str, str]):
    normalized_map = {_normalize_text(k): v for k, v in product_map.items()}
    keys_sorted = sorted(normalized_map.keys(), key=len, reverse=True)

    pattern = re.compile(r"(省心享瞰量[\u4e00-\u9fffA-Za-z0-9\-]+?私募证券投资基金.*?|瞰量[\u4e00-\u9fffA-Za-z0-9\-]+?私募证券投资基金.*?)")

    def match_custodian(file_name: str) -> str | None:
        stem = Path(file_name).stem

        for candidate in pattern.findall(stem):
            key = _normalize_text(candidate)
            if key in normalized_map:
                return normalized_map[key]

        normalized_stem = _normalize_text(stem)
        for product_key in keys_sorted:
            if product_key and product_key in normalized_stem:
                return normalized_map[product_key]
        return None

    return match_custodian


def process_zip_file(input_zip: Path, output_zip: Path) -> dict[str, Any]:
    product_map = _load_product_map()
    match_custodian = _build_matcher(product_map)

    workspace = Path(tempfile.mkdtemp(prefix="invoice_org_"))
    extract_dir = workspace / "extracted"
    output_dir = workspace / "classified"
    extract_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    stats: dict[str, Any] = {
        "input_zip": str(input_zip),
        "output_zip": str(output_zip),
        "extracted_files": 0,
        "pdf_total": 0,
        "copied_total": 0,
        "matched_custodians": 0,
        "unmatched": [],
    }

    try:
        zip_bytes = input_zip.read_bytes()
        stats["extracted_files"] = _extract_zip(zip_bytes, extract_dir)
        pdf_files = _collect_pdfs(extract_dir)
        stats["pdf_total"] = len(pdf_files)

        for pdf_path in pdf_files:
            custodian = match_custodian(pdf_path.name)
            if not custodian:
                stats["unmatched"].append(str(pdf_path.relative_to(extract_dir)))
                continue
            custodian_dir = output_dir / custodian
            custodian_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(pdf_path, custodian_dir / pdf_path.name)
            stats["copied_total"] += 1

        stats["matched_custodians"] = len([p for p in output_dir.iterdir() if p.is_dir()])

        if stats["unmatched"]:
            (output_dir / "_unmatched_files.txt").write_text(
                "\n".join(stats["unmatched"]), encoding="utf-8"
            )

        archive_base = workspace / "invoice_classified"
        archive_path = Path(shutil.make_archive(str(archive_base), "zip", root_dir=output_dir))
        output_zip.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(archive_path, output_zip)
        return stats
    finally:
        shutil.rmtree(workspace, ignore_errors=True)


def run_cli(input_zip: Path, output_zip: Path | None = None) -> int:
    if not input_zip.exists():
        print(f"输入文件不存在: {input_zip}")
        return 1
    if output_zip is None:
        output_zip = input_zip.with_name(f"{input_zip.stem}{DEFAULT_SUFFIX}")

    try:
        stats = process_zip_file(input_zip, output_zip)
    except Exception as exc:
        print(f"处理失败: {exc}")
        return 1

    print("处理完成")
    print(f"输出文件: {stats['output_zip']}")
    print(f"解压文件数: {stats['extracted_files']}")
    print(f"PDF 总数: {stats['pdf_total']}")
    print(f"成功分类: {stats['copied_total']}")
    print(f"未匹配: {len(stats['unmatched'])}")
    return 0


def run_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.title("PDF发票分类工具（本地版）")
    root.geometry("640x300")
    root.resizable(False, False)

    zip_path_var = tk.StringVar()
    out_path_var = tk.StringVar()
    status_var = tk.StringVar(value="请选择输入 ZIP 与输出 ZIP 路径。")

    frame = tk.Frame(root, padx=14, pady=14)
    frame.pack(fill="both", expand=True)

    tk.Label(frame, text="输入 ZIP:").grid(row=0, column=0, sticky="w")
    tk.Entry(frame, textvariable=zip_path_var, width=64).grid(row=1, column=0, sticky="we")

    def pick_input() -> None:
        path = filedialog.askopenfilename(
            title="选择 ZIP 文件", filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
        )
        if path:
            zip_path_var.set(path)
            if not out_path_var.get():
                p = Path(path)
                out_path_var.set(str(p.with_name(f"{p.stem}{DEFAULT_SUFFIX}")))

    tk.Button(frame, text="选择输入", command=pick_input).grid(row=1, column=1, padx=8)

    tk.Label(frame, text="输出 ZIP:").grid(row=2, column=0, sticky="w", pady=(12, 0))
    tk.Entry(frame, textvariable=out_path_var, width=64).grid(row=3, column=0, sticky="we")

    def pick_output() -> None:
        path = filedialog.asksaveasfilename(
            title="保存为",
            defaultextension=".zip",
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")],
        )
        if path:
            out_path_var.set(path)

    tk.Button(frame, text="选择输出", command=pick_output).grid(row=3, column=1, padx=8)

    def run_process() -> None:
        in_text = zip_path_var.get().strip()
        out_text = out_path_var.get().strip()
        if not in_text:
            messagebox.showwarning("提示", "请先选择输入 ZIP 文件。")
            return
        input_zip = Path(in_text)
        output_zip = Path(out_text) if out_text else input_zip.with_name(f"{input_zip.stem}{DEFAULT_SUFFIX}")

        status_var.set("处理中，请稍候...")
        root.update_idletasks()
        try:
            stats = process_zip_file(input_zip, output_zip)
        except Exception as exc:
            status_var.set("处理失败。")
            messagebox.showerror("错误", f"处理失败: {exc}")
            return

        status_var.set("处理完成。")
        summary = (
            f"输出文件: {stats['output_zip']}\n"
            f"解压文件数: {stats['extracted_files']}\n"
            f"PDF 总数: {stats['pdf_total']}\n"
            f"成功分类: {stats['copied_total']}\n"
            f"未匹配: {len(stats['unmatched'])}"
        )
        messagebox.showinfo("完成", summary)

    tk.Button(frame, text="开始分类", command=run_process, width=12).grid(row=4, column=0, sticky="w", pady=14)
    tk.Label(frame, textvariable=status_var, fg="#2f5ea8").grid(row=5, column=0, sticky="w")

    frame.grid_columnconfigure(0, weight=1)
    root.mainloop()


def main() -> None:
    parser = argparse.ArgumentParser(description="PDF 发票分类工具（本地版，无 Streamlit）")
    parser.add_argument("--input", help="输入 ZIP 文件路径")
    parser.add_argument("--output", help="输出 ZIP 文件路径")
    parser.add_argument("--cli", action="store_true", help="使用命令行模式")
    args = parser.parse_args()

    if args.cli:
        if not args.input:
            raise SystemExit("--cli 模式需要提供 --input")
        input_zip = Path(args.input)
        output_zip = Path(args.output) if args.output else None
        raise SystemExit(run_cli(input_zip, output_zip))

    run_gui()


if __name__ == "__main__":
    main()
