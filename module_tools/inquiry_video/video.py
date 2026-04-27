import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.table import _Cell
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import random
from datetime import datetime

# 设置表格边框的函数
def set_cell_border(cell: _Cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    # 添加左边框
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single')
    left.set(qn('w:sz'), '4')
    left.set(qn('w:space'), '0')
    left.set(qn('w:color'), '000000')
    tcPr.append(left)
    
    # 添加右边框
    right = OxmlElement('w:right')
    right.set(qn('w:val'), 'single')
    right.set(qn('w:sz'), '4')
    right.set(qn('w:space'), '0')
    right.set(qn('w:color'), '000000')
    tcPr.append(right)
    
    # 添加上边框
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '4')
    top.set(qn('w:space'), '0')
    top.set(qn('w:color'), '000000')
    tcPr.append(top)
    
    # 添加下边框
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '0')
    bottom.set(qn('w:color'), '000000')
    tcPr.append(bottom)

# 读取Excel文件
df = pd.read_excel('video.xlsx')

# 处理每一行数据
for index, row in df.iterrows():
    # 获取数据
    code = str(row['代码']).zfill(6)  # 补齐六位数
    name = row['名称']
    inquiry_date = row['询价日']
    
    # 格式化日期
    if isinstance(inquiry_date, datetime):
        date_str = inquiry_date.strftime('%Y年%m月%d日')
    else:
        # 尝试从字符串解析日期
        try:
            date_obj = datetime.strptime(str(inquiry_date), '%Y-%m-%d')
            date_str = date_obj.strftime('%Y年%m月%d日')
        except:
            date_str = str(inquiry_date)
    
    # 创建Word文档
    doc = Document()
    
    # 设置全局字体和行距
    style = doc.styles['Normal']
    style.font.name = '仿宋'
    # 设置中文字体
    style.font.element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
    style.font.size = Pt(15)  # 小三字体
    style.paragraph_format.line_spacing = 1.5
    
    # 添加标题（正文格式，居中）
    title = doc.add_paragraph('新股询价现场通讯工具上交登记表')
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    title_run = title.runs[0]
    title_run.font.name = '仿宋'
    # 设置中文字体
    title_run.font.element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
    title_run.font.size = Pt(15)
    
    # 添加内容
    doc.add_paragraph(f'股票名称：【{code}】')
    doc.add_paragraph(f'股票代码：【{name}】')
    doc.add_paragraph(f'询价日期：{date_str}')
    doc.add_paragraph('询价关键时间窗口：09：30 - 15：00')
    doc.add_paragraph('保管地点：【合规部】')
    doc.add_paragraph()  # 空行
    
    # 创建表格
    table = doc.add_table(rows=5, cols=5)
    
    # 设置表格标题行
    headers = ['人员', '上交时间', '通讯工具类型', '上交确认', '领取时间']
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = header
        # 设置表格文字字体
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.name = '仿宋'
                # 设置中文字体
                run.font.element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
                run.font.size = Pt(15)
        # 设置表格边框
        set_cell_border(cell)
    
    # 人员列表
    persons = ['韩震泓', '余晓舰', '王博', '罗钰瑶']
    
    # 填充表格数据
    for i, person in enumerate(persons, start=1):
        # 生成随机上交时间（9:00-9:29）
        hour = 9
        minute = random.randint(0, 29)
        hand_in_time = f'{hour:02d}：{minute:02d}'
        
        # 生成随机领取时间（15:01-15:30）
        hour = 15
        minute = random.randint(1, 30)
        pick_up_time = f'{hour:02d}：{minute:02d}'
        
        # 填充表格
        cells = [
            table.cell(i, 0),
            table.cell(i, 1),
            table.cell(i, 2),
            table.cell(i, 3),
            table.cell(i, 4)
        ]
        values = [person, hand_in_time, '手机', '√', pick_up_time]
        
        for cell, value in zip(cells, values):
            cell.text = value
            # 设置表格文字字体
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = '仿宋'
                    run.font.size = Pt(15)
            # 设置表格边框
            set_cell_border(cell)
    
    # 格式化文件名中的日期
    if isinstance(inquiry_date, datetime):
        date_file = inquiry_date.strftime('%Y%m%d')
    else:
        try:
            date_obj = datetime.strptime(str(inquiry_date), '%Y-%m-%d')
            date_file = date_obj.strftime('%Y%m%d')
        except:
            date_file = str(inquiry_date).replace('-', '')
    
    # 保存文件
    filename = f'{code}_{name}_{date_file}_通讯工具上交登记表.docx'
    doc.save(filename)
    print(f'已生成文件：{filename}')

print('所有文件生成完成！')
