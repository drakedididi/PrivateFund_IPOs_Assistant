import os
import re
import openpyxl
from tqdm import tqdm

# 本地存放txt文件的相对路径
txt_folder1 = 'etf_txts深1'
txt_folder2 = 'etf_txts深2'
output_excel = 'etf_full_result.xlsx'

def extract_from_txt(txt_path, target_code):
    with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()

    etf_code = os.path.splitext(os.path.basename(txt_path))[0]
    stock_section = False
    records = []

    for line in lines:
        if '组合信息内容' in line:
            stock_section = True
            continue
        if stock_section:
            if not line.strip():
                break
            parts = line.strip().split()
            if len(parts) >= 2:
                stock_code = parts[0]
                # 股票代码支持5~6位数字
                if re.fullmatch(r'\d{5,6}', stock_code):
                    flag = None
                    for part in parts:
                        if part in ('允许', '必须'):
                            flag = part
                            break
                    if stock_code == target_code and flag:
                        records.append((etf_code, stock_code, flag))
    return records

def parse_all_txts(target_code):
    all_records = []
    for folder in [txt_folder1, txt_folder2]:
        for fname in tqdm(os.listdir(folder), desc=f'解析TXT文件 - {folder}'):
            if fname.endswith('.txt'):
                path = os.path.join(folder, fname)
                all_records.extend(extract_from_txt(path, target_code))
    return all_records

def save_to_excel(records, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['ETF代码', '股票代码', '现金替代标志'])
    for row in records:
        ws.append(row)
    wb.save(filename)

if __name__ == '__main__':
    target_code = input("请输入股票代码：").strip()
    all_data = parse_all_txts(target_code)
    save_to_excel(all_data, output_excel)
    print(f"[INFO] 已保存至 {output_excel}")