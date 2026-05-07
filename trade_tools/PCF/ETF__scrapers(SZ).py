

import asyncio
import os
from playwright.async_api import async_playwright
import openpyxl
import time


def get_playwright_headless(default: bool = False) -> bool:
    raw = os.getenv("PLAYWRIGHT_HEADLESS")
    if raw is None:
        return default
    value = raw.strip().lower()
    if not value:
        return default
    return value in {"1", "true", "yes", "on"}

async def scrape_etf_links():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['ETF代码', '下载链接'])

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=get_playwright_headless())
        page = await browser.new_page()
        await page.goto("https://www.szse.cn/disclosure/fund/currency/index.html", wait_until="load")
        await page.wait_for_timeout(3000)

        while True:
            rows = await page.query_selector_all("table tbody tr")
            for row in rows:
                a_tags = await row.query_selector_all("td a")
                if len(a_tags) >= 2:
                    etf_code = (await a_tags[0].inner_text()).strip()
                    href = await a_tags[1].get_attribute('href')
                    if href:
                        full_link = 'https://www.szse.cn' + href if href.startswith('/') else href
                        ws.append([etf_code, full_link])

            next_button_li = await page.query_selector('li.next')
            if next_button_li:
                class_attr = await next_button_li.get_attribute('class') or ''
                if 'disabled' in class_attr:
                    break
                next_a = await next_button_li.query_selector('a')
                if next_a:
                    await next_a.click()
                    await page.wait_for_timeout(2000)
                else:
                    break
            else:
                break

        await browser.close()
    wb.save("etf_links(SZ).xlsx")

if __name__ == '__main__':
    asyncio.run(scrape_etf_links())
