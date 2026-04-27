import asyncio
import os
import re
import requests
import time
from tqdm import tqdm
from openpyxl import load_workbook
from playwright.async_api import async_playwright

link_excel = 'etf_links(SH).xlsx'
txt_folder1 = 'etf_txts沪1'
txt_folder2 = 'etf_txts沪2'
os.makedirs(txt_folder1, exist_ok=True)
os.makedirs(txt_folder2, exist_ok=True)

subtype_button_xpath_template = '//*[@id="tab_main"]/div[1]/div[1]/button[{}]'
# 表格字段 xpath
row_xpath_base = '//*[@id="tab_main"]/div[1]/div[2]/table/tbody/tr'
code_xpath = ':scope > td:nth-child(1)'
download_xpath = ':scope > td:nth-child(7) a'

subtype_pages = {
    1: 2,
    2: 16,
    3: 2,
    4: 3,
    5: 1
}

def clean_etf_code(raw):
    match = re.search(r'(ETF)?(\d{6})', str(raw))
    return match.group(2) if match else str(raw)

async def download_txt_files():
    wb = load_workbook(link_excel)
    ws = wb.active
    success_count = 0

    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
        'Referer': 'https://www.sse.com.cn/disclosure/fund/etflist/',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    }

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.goto('https://www.sse.com.cn/disclosure/fund/etflist/')
        await page.wait_for_selector('#tab_main')
        
        for subtype_index in range(1, 6):  # 遍历5个类别
            print(f"\n===== 开始处理第 {subtype_index} 类 ETF =====")
            button_xpath = subtype_button_xpath_template.format(subtype_index)
            await page.click(button_xpath)
            await asyncio.sleep(1.5)  # 增加等待时间确保加载完成
            
            current_page = 1
            total_pages = subtype_pages[subtype_index]
            
            while current_page <= total_pages:
                print(f"正在处理第 {current_page}/{total_pages} 页...")
                await page.wait_for_selector('a.js_download-export', timeout=10000)
                
                # 获取当前页所有下载链接
                links = await page.query_selector_all('a.js_download-export')
                print(f"本页找到 {len(links)} 个下载链接")
                
                # 处理当前页所有链接
                for idx, link in enumerate(links):
                    try:
                        href = await link.get_attribute('href')
                        if not href:
                            print(f"链接 {idx} 没有 href 属性")
                            continue
                            
                        # 处理URL
                        if href.startswith('http'):
                            txt_url = href
                        elif href.startswith('//'):
                            txt_url = 'https:' + href
                        elif href.startswith('/query.sse.com.cn'):
                            txt_url = 'https:/' + href  # 处理特殊格式
                        else:
                            href = href.lstrip('/')
                            txt_url = f'https://query.sse.com.cn/{href}'
                            
                        # 获取ETF代码
                        row = await link.query_selector('xpath=../..')  # 向上两级到tr
                        if row:
                            code_element = await row.query_selector(code_xpath)
                            if code_element:
                                etf_code_raw = await code_element.text_content()
                                etf_code = clean_etf_code(etf_code_raw)
                            else:
                                print("无法获取代码元素")
                                continue
                        else:
                            print("无法获取行元素")
                            continue
                            
                        print(f"正在下载 {etf_code} 的TXT文件: {txt_url}")
                        resp = requests.get(txt_url, headers=headers, timeout=15)
                        
                        if resp.status_code == 200:
                            # 检查是否是HTML文件
                            if 'html' in resp.headers.get('Content-Type', '').lower() or \
                               '<html' in resp.text[:100].lower():
                                print(f'[WARN] 返回HTML内容: {etf_code} {txt_url}')
                                continue
                                
                            folder = txt_folder1 if success_count < 300 else txt_folder2
                            save_path = os.path.join(folder, f'{etf_code}.txt')
                            
                            with open(save_path, 'wb') as f:
                                f.write(resp.content)
                            
                            success_count += 1
                            print(f"第{success_count}个ETF {etf_code} 下载成功")
                            time.sleep(0.9)  # 适当降低等待时间
                        else:
                            print(f'[WARN] 下载失败: HTTP {resp.status_code} {etf_code} {txt_url}')
                    except Exception as e:
                        print(f'[ERROR] 下载失败: {etf_code if "etf_code" in locals() else "未知代码"} - {e}')
                
                # 翻页处理
                if current_page < total_pages:
                    try:
                        print("尝试翻到下一页...")
                        next_buttons = await page.query_selector_all('li.next')
                        if next_buttons:
                            next_button = next_buttons[0]
                            class_attr = await next_button.get_attribute('class') or ''
                            
                            if 'disabled' in class_attr:
                                print("已经是最后一页，停止翻页")
                                break
                            
                            await next_button.click()
                            await page.wait_for_selector(row_xpath_base, timeout=10000)
                            await asyncio.sleep(1.5)  # 等待新页面加载
                            current_page += 1
                        else:
                            print("未找到下一页按钮")
                            break
                    except Exception as e:
                        print(f"翻页出错: {e}")
                        break
                else:
                    break  # 所有页面处理完成
                    
        await browser.close()
        print(f"\n所有类别处理完成! 共成功下载 {success_count} 个文件")

if __name__ == '__main__':
    try:
        # 尝试获取当前运行的事件循环
        loop = asyncio.get_event_loop()
        
        # 如果事件循环已经在运行（如在Spyder中）
        if loop.is_running():
            # 在现有事件循环中运行
            loop.create_task(download_txt_files())
            print("任务已在现有事件循环中启动")
        else:
            # 否则创建新的事件循环
            loop.run_until_complete(download_txt_files())
    except RuntimeError:
        # 如果没有事件循环，创建新的事件循环
        asyncio.run(download_txt_files())