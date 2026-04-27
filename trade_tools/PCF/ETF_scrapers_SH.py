import asyncio
from playwright.async_api import async_playwright
import openpyxl

# 小类按钮xpath（index从1到5）
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

# 表格字段 xpath
row_xpath_base = '//*[@id="tab_main"]/div[1]/div[2]/table/tbody/tr'
code_xpath = ':scope > td:nth-child(1)'
download_xpath = ':scope > td:nth-child(7) a'

async def scrape_etf():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['ETF代码', '下载链接'])

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.goto('https://www.sse.com.cn/disclosure/fund/etflist/')

        await page.wait_for_selector('#tab_main')

        for subtype_index in range(1, 6):
            
            button_xpath = subtype_button_xpath_template.format(subtype_index)
            await page.click(button_xpath)
            await asyncio.sleep(1)
            print(f"进入第{subtype_index}类ETF")
            total_pages = subtype_pages[subtype_index]
            for page_index in range(total_pages):
                print(f"进入第{page_index+1}页")
                await page.wait_for_selector(row_xpath_base)
                rows = await page.query_selector_all(row_xpath_base)
                for row in rows:
                    code_elem = await row.query_selector(code_xpath)
                    link_elem = await row.query_selector(download_xpath)
                    if code_elem and link_elem:
                        code = (await code_elem.inner_text()).strip()
                        href = await link_elem.get_attribute('href')
                        if code and href:
                            full_link = 'https://www.sse.com.cn' + href if href.startswith('/') else href
                            ws.append([code, full_link])
                            print(f"{code}的链接已保存")
                if page_index < total_pages - 1:
                    try:
                        # 查找下一页按钮
                        next_buttons = await page.query_selector_all('li.next')
                        if not next_buttons:
                            print("    未找到下一页按钮")
                            continue
                        
                        next_button = next_buttons[0]
                        class_attr = await next_button.get_attribute('class') or ''
                        
                        # 检查是否已禁用（最后一页）
                        if 'disabled' in class_attr:
                            print("    已是最后一页")
                            break
                        
                        # 点击下一页
                        next_link = await next_button.query_selector('a')
                        if next_link:
                            await next_link.click()
                            # 等待页面加载完成
                            await page.wait_for_selector(row_xpath_base)
                            await asyncio.sleep(1)
                        else:
                            print("    未找到下一页链接")
                            break
                            
                    except Exception as e:
                        print(f"    翻页出错: {e}")
                        break
                    
        wb.save('etf_links(SH).xlsx')
        print("[SUCCESS] 已保存为 etf_links(SH).xlsx")
        await browser.close()

if __name__ == '__main__':
    try:
        # 尝试获取当前运行的事件循环
        loop = asyncio.get_event_loop()
        
        # 如果事件循环已经在运行（如在Spyder中）
        if loop.is_running():
            # 在现有事件循环中运行
            loop.create_task(scrape_etf())
            print("任务已在现有事件循环中启动")
        else:
            # 否则创建新的事件循环
            loop.run_until_complete(scrape_etf())
    except RuntimeError:
        # 如果没有事件循环，创建新的事件循环
        asyncio.run(scrape_etf())