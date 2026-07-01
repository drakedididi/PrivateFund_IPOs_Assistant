from __future__ import annotations

import argparse
import csv
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
import xlrd


PRODUCT_COLUMNS = ("产品", "产品名")
DEFAULT_SUFFIX = "_pdf_withdraw.zip"


def normalize_text(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    return re.sub(r"[\s\-_()（）【】\[\]《》<>.,，。·•、:：;；/\\]+", "", text).lower()


def safe_filename(value: str, fallback: str = "product") -> str:
    name = re.sub(r'[<>:"/\\|?*\r\n\t]+', "_", value).strip(" ._")
    return name[:120] or fallback


def _decode_zip_name(name: str) -> str:
    try:
        repaired = name.encode("cp437").decode("gbk")
    except Exception:
        return name
    old_cjk = sum(1 for ch in name if "\u4e00" <= ch <= "\u9fff")
    new_cjk = sum(1 for ch in repaired if "\u4e00" <= ch <= "\u9fff")
    return repaired if new_cjk > old_cjk else name


def _safe_target_path(base_dir: Path, member_name: str) -> Path:
    target = (base_dir / Path(member_name)).resolve()
    base = base_dir.resolve()
    if base != target and base not in target.parents:
        raise ValueError(f"非法路径: {member_name}")
    return target


def extract_zip(input_zip: Path, extract_dir: Path) -> int:
    count = 0
    with zipfile.ZipFile(input_zip) as zf:
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


def _find_header_index(values: list[Any]) -> int | None:
    for index, value in enumerate(values):
        text = "" if value is None else str(value).strip()
        if text in PRODUCT_COLUMNS:
            return index
    return None


def _dedupe_products(products: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for product in products:
        text = str(product).strip()
        key = normalize_text(text)
        if not text or not key or key in seen:
            continue
        seen.add(key)
        result.append(text)
    return result


def read_products_from_excel(excel_path: Path) -> list[str]:
    suffix = excel_path.suffix.lower()
    if suffix == ".xlsx":
        return _read_products_from_xlsx(excel_path)
    if suffix == ".xls":
        return _read_products_from_xls(excel_path)
    raise ValueError("仅支持 XLSX 或 XLS 文件")


def _read_products_from_xlsx(excel_path: Path) -> list[str]:
    workbook = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    try:
        for sheet in workbook.worksheets:
            rows = sheet.iter_rows(values_only=True)
            for row_number, row in enumerate(rows, start=1):
                if row_number > 30:
                    break
                column_index = _find_header_index(list(row))
                if column_index is None:
                    continue
                products = [
                    "" if row_values[column_index] is None else str(row_values[column_index]).strip()
                    for row_values in rows
                    if column_index < len(row_values)
                ]
                return _dedupe_products(products)
    finally:
        workbook.close()
    raise ValueError("未找到列名为“产品”或“产品名”的列")


def _read_products_from_xls(excel_path: Path) -> list[str]:
    workbook = xlrd.open_workbook(str(excel_path))
    for sheet in workbook.sheets():
        for row_number in range(min(sheet.nrows, 30)):
            values = sheet.row_values(row_number)
            column_index = _find_header_index(values)
            if column_index is None:
                continue
            products = [
                str(sheet.cell_value(index, column_index)).strip()
                for index in range(row_number + 1, sheet.nrows)
                if column_index < sheet.ncols
            ]
            return _dedupe_products(products)
    raise ValueError("未找到列名为“产品”或“产品名”的列")


def collect_pdfs(root_dir: Path) -> list[Path]:
    return sorted(p for p in root_dir.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf")


def unique_destination(base_dir: Path, filename: str) -> Path:
    target = base_dir / filename
    if not target.exists():
        return target
    stem = target.stem
    suffix = target.suffix
    index = 2
    while True:
        candidate = base_dir / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def match_pdfs_by_products(products: list[str], pdf_files: list[Path]) -> dict[str, list[Path]]:
    normalized_files = [(pdf_path, normalize_text(pdf_path.stem)) for pdf_path in pdf_files]
    matches: dict[str, list[Path]] = {}
    for product in _dedupe_products(products):
        product_key = normalize_text(product)
        if product_key:
            matches[product] = [pdf_path for pdf_path, pdf_key in normalized_files if product_key in pdf_key]
    return matches


def withdraw_pdfs(input_zip: Path, products: list[str], output_zip: Path) -> dict[str, Any]:
    if not products:
        raise ValueError("产品列表为空，请先读取 Excel")

    workspace = Path(tempfile.mkdtemp(prefix="pdf_withdraw_"))
    extract_dir = workspace / "extracted"
    output_dir = workspace / "selected_pdfs"
    extract_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        extracted_count = extract_zip(input_zip, extract_dir)
        pdf_files = collect_pdfs(extract_dir)
        matches = match_pdfs_by_products(products, pdf_files)

        copied_total = 0
        matched_products = 0
        report_rows: list[dict[str, str]] = []
        unmatched_products: list[str] = []

        for product, matched_files in matches.items():
            if not matched_files:
                unmatched_products.append(product)
                report_rows.append({"product": product, "pdf": "", "status": "unmatched"})
                continue

            matched_products += 1
            product_dir = output_dir / safe_filename(product)
            product_dir.mkdir(parents=True, exist_ok=True)
            for pdf_path in matched_files:
                destination = unique_destination(product_dir, pdf_path.name)
                shutil.copy2(pdf_path, destination)
                copied_total += 1
                report_rows.append(
                    {
                        "product": product,
                        "pdf": str(pdf_path.relative_to(extract_dir)),
                        "status": "matched",
                    }
                )

        with (output_dir / "_match_report.csv").open("w", newline="", encoding="utf-8-sig") as report_file:
            writer = csv.DictWriter(report_file, fieldnames=["product", "pdf", "status"])
            writer.writeheader()
            writer.writerows(report_rows)

        if unmatched_products:
            (output_dir / "_unmatched_products.txt").write_text(
                "\n".join(unmatched_products), encoding="utf-8"
            )

        archive_base = workspace / "pdf_withdraw"
        archive_path = Path(shutil.make_archive(str(archive_base), "zip", root_dir=output_dir))
        output_zip.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(archive_path, output_zip)

        return {
            "extracted_files": extracted_count,
            "pdf_total": len(pdf_files),
            "product_total": len(matches),
            "matched_products": matched_products,
            "copied_total": copied_total,
            "unmatched_products": len(unmatched_products),
            "output_zip": str(output_zip),
        }
    finally:
        shutil.rmtree(workspace, ignore_errors=True)


def read_products_from_excel_bytes(file_bytes: bytes, suffix: str) -> list[str]:
    with tempfile.TemporaryDirectory(prefix="pdf_withdraw_excel_") as tmp_dir:
        excel_path = Path(tmp_dir) / f"products{suffix.lower()}"
        excel_path.write_bytes(file_bytes)
        return read_products_from_excel(excel_path)


def run_cli(excel_path: Path, input_zip: Path, output_zip: Path | None = None) -> int:
    if output_zip is None:
        output_zip = input_zip.with_name(f"{input_zip.stem}{DEFAULT_SUFFIX}")
    products = read_products_from_excel(excel_path)
    stats = withdraw_pdfs(input_zip, products, output_zip)
    print(f"读取产品数量: {stats['product_total']}")
    print(f"ZIP 内 PDF 数量: {stats['pdf_total']}")
    print(f"匹配产品数量: {stats['matched_products']}")
    print(f"复制 PDF 数量: {stats['copied_total']}")
    print(f"输出文件: {stats['output_zip']}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="根据 Excel 产品名单从 ZIP 中抽取匹配 PDF")
    parser.add_argument("excel", type=Path, help="包含“产品”或“产品名”列的 Excel")
    parser.add_argument("zip", type=Path, help="包含 PDF 文件夹的 ZIP")
    parser.add_argument("-o", "--output", type=Path, help="输出 ZIP 路径")
    args = parser.parse_args()
    return run_cli(args.excel, args.zip, args.output)


if __name__ == "__main__":
    raise SystemExit(main())
