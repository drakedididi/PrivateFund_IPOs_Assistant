import os
import re
import openpyxl
from tqdm import tqdm

# 本地存放txt文件的相对路径
txt_folder1 = 'etf_txts沪1'
txt_folder2 = 'etf_txts沪2'
output_excel = 'etf_full_result.xlsx'

def extract_from_txt(txt_path, target_code):
    with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()

    etf_code = os.path.splitext(os.path.basename(txt_path))[0]
    stock_section = False
    records = []

    for line in lines:
        if 'TAGTAG' in line:  # 改为识别 TAGTAG 标记
            stock_section = True
            continue
        if stock_section:
            if 'ENDENDEND' in line or not line.strip():  # 增加 ENDENDEND 结束标记
                break
            parts = [p.strip() for p in line.strip().split('|')]  # 使用 | 作为分隔符
            if len(parts) >= 4 and parts[0]:  # 至少需要4个字段且股票代码不为空
                stock_code = parts[0]
                # 股票代码支持5~6位数字
                if re.fullmatch(r'\d{5,6}', stock_code):
                    cash_flag = parts[3]  # 第4个字段是现金替代标志
                    # 检查现金替代标志是否为0、1或2
                    if cash_flag in ('0', '1', '2'):
                        if stock_code == target_code:
                            records.append((etf_code, stock_code, cash_flag))
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