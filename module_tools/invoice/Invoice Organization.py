import os
import re
import shutil
from pathlib import Path


PRODUCT_TO_CUSTODIAN = {
    "睿量电子20号私募证券投资基金": "招商证券股份有限公司",
    "睿量信选原子3号私募证券投资基金": "中信证券股份有限公司",
    "睿量HT鹏泰1号私募证券投资基金": "华泰证券股份有限公司",
    "睿量量子-江海远山51期私募证券投资基金": "招商证券股份有限公司",
    "睿量信淮原子9号1期私募证券投资基金": "中信建投证券股份有限公司",
    "睿量信淮原子9号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量信淮原子11号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量原子7号私募证券投资基金2期": "招商证券股份有限公司",
    "睿量原子7号私募证券投资基金": "招商证券股份有限公司",
    "睿量原子7号私募证券投资基金1期": "招商证券股份有限公司",
    "睿量量子聚利1号1期私募证券投资基金": "中泰证券股份有限公司",
    "睿量量子春晓3号私募证券投资基金": "国泰君安证券股份有限公司",
    "睿量陆享原子16号私募证券投资基金": "招商证券股份有限公司",
    "睿量原子8号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量原子2号二期私募证券投资基金": "国信证券股份有限公司",
    "睿量陆享原子6号私募证券投资基金": "招商证券股份有限公司",
    "睿量中远1号私募证券投资基金": "华泰证券股份有限公司",
    "睿量原子2号私募证券投资基金": "国信证券股份有限公司",
    "睿量原子2号一期私募证券投资基金": "国信证券股份有限公司",
    "睿量原子5号私募证券投资基金": "长江证券股份有限公司",
    "睿量原子10号私募证券投资基金": "广发证券股份有限公司",
    "睿量电子22号私募证券投资基金": "华泰证券股份有限公司",
    "睿量量子聚利1号私募证券投资基金": "中泰证券股份有限公司",
    "睿量电子17号私募证券投资基金": "华泰证券股份有限公司",
    "睿量量子15号私募证券投资基金": "广发证券股份有限公司",
    "睿量臻选1000指数增强1号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量质子1号私募证券投资基金": "招商证券股份有限公司",
    "睿量中微子2号私募证券投资基金": "国泰君安证券股份有限公司",
    "省心享睿量兴泰奋进1期私募证券投资基金": "华泰证券股份有限公司",
    "省心享睿量兴泰锐进1号私募证券投资基金": "华泰证券股份有限公司",
    "睿量电子18号私募证券投资基金": "华泰证券股份有限公司",
    "睿量电子21号私募证券投资基金": "招商证券股份有限公司",
    "睿量信淮智行9号1期私募证券投资基金": "中信建投证券股份有限公司",
    "睿量信淮智行9号2期私募证券投资基金": "中信建投证券股份有限公司",
    "睿量量子12号私募证券投资基金": "国信证券股份有限公司",
    "睿量量子13号私募证券投资基金": "国信证券股份有限公司",
    "睿量信淮智行9号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量智行7号私募证券投资基金2期": "招商证券股份有限公司",
    "睿量光子6号二期私募证券投资基金": "国信证券股份有限公司",
    "睿量光子6号私募证券投资基金": "国信证券股份有限公司",
    "睿量红光1号私募证券投资基金": "华泰证券股份有限公司",
    "睿量光子6号一期私募证券投资基金": "国信证券股份有限公司",
    "睿量智行7号私募证券投资基金1期": "招商证券股份有限公司",
    "睿量智享2号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量智行7号私募证券投资基金": "招商证券股份有限公司",
    "睿量电子15号私募证券投资基金": "招商证券股份有限公司",
    "睿量智行广睿1号私募证券投资基金": "广发证券股份有限公司",
    "睿量光子1号1期私募证券投资基金": "华泰证券股份有限公司",
    "睿量智行1号1期私募证券投资基金": "华泰证券股份有限公司",
    "睿量光子春晓2号私募证券投资基金": "国泰君安证券股份有限公司",
    "睿量电子9号私募证券投资基金": "华泰证券股份有限公司",
    "睿量福兴2号私募证券投资基金": "兴业证券股份有限公司",
    "睿量电子12号私募证券投资基金": "国信证券股份有限公司",
    "睿量电子13号私募证券投资基金": "国信证券股份有限公司",
    "睿量中子3号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量电子5号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量量子3号私募证券投资基金": "兴业证券股份有限公司",
    "睿量量子11号私募证券投资基金": "国信证券股份有限公司",
    "睿量电子7号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量量子春晓1号私募证券投资基金": "国泰君安证券股份有限公司",
    "睿量电子6号私募证券投资基金": "国泰君安证券股份有限公司",
    "睿量红星2号私募证券投资基金": "华泰证券股份有限公司",
    "睿量电子3号私募证券投资基金": "华泰证券股份有限公司",
    "睿量电子2号私募证券投资基金": "华泰证券股份有限公司",
    "睿量电子11号私募证券投资基金": "华泰证券股份有限公司",
    "睿量量子9号私募证券投资基金": "华泰证券股份有限公司",
    "睿量中子2号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量智行3号私募证券投资基金": "华泰证券股份有限公司",
    "睿量量子10号私募证券投资基金": "海通证券股份有限公司",
    "睿量智行2号私募证券投资基金": "华泰证券股份有限公司",
    "睿量光子1号私募证券投资基金": "华泰证券股份有限公司",
    "睿量量子1号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量电子1号私募证券投资基金": "华泰证券股份有限公司",
    "睿量红星1号私募证券投资基金": "华泰证券股份有限公司",
    "睿量中微子1号私募证券投资基金": "华泰证券股份有限公司",
    "睿量分子1号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量中子1号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量原子1号私募证券投资基金": "中信建投证券股份有限公司",
    "睿量智行1号私募证券投资基金": "华泰证券股份有限公司",
    "睿量智享1号私募证券投资基金": "中信建投证券股份有限公司",
}


def normalize_text(text: str) -> str:
    """Normalize names to improve matching reliability."""
    text = text.strip()
    text = re.sub(r"[\s\-—_（）()【】\[\]·,，。]", "", text)
    return text


NORMALIZED_PRODUCT_TO_CUSTODIAN = {
    normalize_text(product): custodian
    for product, custodian in PRODUCT_TO_CUSTODIAN.items()
}

PRODUCT_KEYS_SORTED = sorted(
    NORMALIZED_PRODUCT_TO_CUSTODIAN.keys(), key=len, reverse=True
)

# Prefer extracting candidate product names such as "睿量xxx私募证券投资基金" from filenames.
PRODUCT_NAME_PATTERN = re.compile(
    r"(省心享睿量[\u4e00-\u9fffA-Za-z0-9\-]+?私募证券投资基金(?:[0-9一二三四五六七八九十]+期)?|"
    r"睿量[\u4e00-\u9fffA-Za-z0-9\-]+?私募证券投资基金(?:[0-9一二三四五六七八九十]+期)?)"
)


def match_custodian(pdf_name: str) -> str | None:
    stem = Path(pdf_name).stem

    # First pass: regex extraction from original filename.
    for candidate in PRODUCT_NAME_PATTERN.findall(stem):
        candidate_key = normalize_text(candidate)
        if candidate_key in NORMALIZED_PRODUCT_TO_CUSTODIAN:
            return NORMALIZED_PRODUCT_TO_CUSTODIAN[candidate_key]

    # Fallback: longest known product name contained in normalized filename.
    normalized_stem = normalize_text(stem)
    for product_key in PRODUCT_KEYS_SORTED:
        if product_key in normalized_stem:
            return NORMALIZED_PRODUCT_TO_CUSTODIAN[product_key]

    return None


def copy_invoice_pdfs() -> None:
    script_dir = Path(__file__).resolve().parent
    custodian_names = set(PRODUCT_TO_CUSTODIAN.values())

    copied_count = 0
    unmatched_files = []

    for root, dirs, files in os.walk(script_dir, topdown=True):
        # Avoid recursively processing destination folders.
        dirs[:] = [
            d for d in dirs
            if d not in custodian_names and not d.startswith(".")
        ]

        for file_name in files:
            if not file_name.lower().endswith(".pdf"):
                continue

            source_path = Path(root) / file_name
            custodian = match_custodian(file_name)
            if not custodian:
                try:
                    rel_path = source_path.relative_to(script_dir)
                except ValueError:
                    rel_path = source_path
                unmatched_files.append(str(rel_path))
                continue

            target_dir = script_dir / custodian
            target_dir.mkdir(parents=True, exist_ok=True)

            target_path = target_dir / file_name

            # If file is already in the destination folder, skip it.
            if source_path.resolve() == target_path.resolve():
                continue

            shutil.copy2(source_path, target_path)
            copied_count += 1

    print(f"一共移动了 {copied_count} 个PDF")
    if unmatched_files:
        print(f"未匹配托管名称的PDF数量: {len(unmatched_files)}")
        for item in unmatched_files[:20]:
            print(f"  - {item}")


if __name__ == "__main__":
    copy_invoice_pdfs()
