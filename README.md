# PrivateFund_IPOs_Assistant

本链接 (https://fundassistant.pages.dev/frontend/IPO) 仅展示部分线上功能，如有个性化需求，部署【本地版】请联系管理员

PrivateFund_IPOs_Assistant 是一个面向私募基金 打新/运营/交易场景的 Web 化工具集合。项目以静态前端页面承载业务入口，以 Flask 提供文件处理、数据爬取和交易辅助 API，可部署为 Cloudflare Pages 前端 + Render 后端，也支持部分交易工具在本地启动使用。

## 功能概览

### 日历

日历页面位于 `frontend/IPO.html`，用于集中展示打新相关时间安排和辅助分析信息。

- A 股日历：展示沪深京新股申购、缴款、上市、限售解禁等日期。
- 可转债日历：展示可转债申购、缴款、上市安排，并提供双低值区域。
- H 股日历：展示 H 股新股招股、申购节奏和上市安排。
- 数据来源：通过 `crawlers/` 下的交易所、东方财富、可转债等爬虫抓取原始数据，再由 `merge/` 脚本合并为前端可直接读取的 JSON。
- 自动更新：`.github/workflows/daily_crawl.yml` 会在 GitHub Actions 中定时运行爬虫、同步前端数据，并部署静态页面到 Cloudflare Pages。
- 分享辅助：前端支持将日历视图生成图片，便于业务沟通和日报/周报复用。

核心文件：

| 路径 | 说明 |
| --- | --- |
| `frontend/IPO.html` | IPO/可转债/H 股日历主页面 |
| `frontend/unlock_db.js` | 限售解禁本地数据补充 |
| `crawlers/` | 上交所、深交所、北交所、东方财富、可转债数据爬虫 |
| `merge/` | A 股、H 股、综合日历数据合并 |
| `data/trading_holidays.json` | 交易日/节假日基础数据 |
| `scramer_merge.py` | GitHub Actions 调用的总入口脚本 |
| `sync_unlock_db.py` | 同步限售解禁数据到前端脚本 |

### 运营组件

运营组件页面位于 `frontend/module_tools.html`，通过文件上传调用 Flask API，处理完成后返回 ZIP、Excel、Word 或图片结果。部分涉及内部数据的接口需要 `APP_SECRET_TOKEN` 鉴权。

| 功能 | API | 输入 | 输出 | 说明 |
| --- | --- | --- | --- | --- |
| 发票分类 | `POST /api/invoice` | ZIP | 分类后的 ZIP | 按 PDF 文件名中的托管人/产品信息自动归类妥妥递发票 |
| 新股询价登记表 | `POST /api/video` | ZIP/XLSX/XLS | Word 文档 ZIP | 根据代码、名称、询价日期批量生成上交登记表 |
| 超额收益回撤图 | `POST /api/extra-revenue` | XLSX/XLS | PNG 图片 ZIP | 自动识别日频/周频，生成产品 vs 指数超额收益回撤及修复周期图 |
| 巨额赎回公告 | `POST /api/redemption` | ZIP/XLSX/XLS | Word 文档 ZIP | 批量生成触发巨额赎回公告 |
| 估值表资产提取 | `POST /api/valuation-table` | PDF/ZIP | Excel ZIP | 从估值表 PDF 中提取打新产品总资产，并标记低于阈值的产品 |
| PDF 抽取 | `POST /api/pdf-withdraw/products`、`POST /api/pdf-withdraw` | Excel + PDF ZIP | 匹配后的 PDF ZIP | 先读取产品清单，再从 PDF 压缩包中按产品名抽取相关文件 |
| 关联交易决策留档 | `POST /api/excel` | ZIP/XLSX/XLS | Word 文档 ZIP | 根据交易确认明细批量生成关联交易决策机制留档 |
| 关联交易公告 | `POST /api/fund` | ZIP/XLSX/XLS | Word 文档 ZIP | 根据交易确认明细批量生成触发关联交易公告 |

核心文件：

| 路径 | 说明 |
| --- | --- |
| `app.py` | Flask 后端入口，统一处理上传校验、临时目录、安全解压、工具调用和结果下载 |
| `frontend/module_tools.html` | 运营组件前端页面 |
| `module_tools/invoice/` | 发票分类逻辑 |
| `module_tools/inquiry_video/video.py` | 新股询价登记表生成 |
| `module_tools/extra_revenue/extra_revenue.py` | 超额收益回撤图生成 |
| `module_tools/big_redemption/redemption_word.py` | 巨额赎回公告生成 |
| `module_tools/valuation_table/valuation_table.py` | 估值表 PDF 解析 |
| `module_tools/withdraw.py` | 产品 PDF 抽取 |
| `module_tools/related_deal/` | 关联交易决策留档和公告生成 |

### 交易组件

交易组件页面位于 `frontend/trade_tools.html`，当前核心功能是 ETF PCF 白名单查询和 PCF 成分导出。该组件更适合在本地运行，因为 PCF 数据会落在本机持久目录中。

- 支持从上交所、深交所刷新 ETF PCF 下载链接和 PCF 文本文件。
- 输入股票代码后，查询包含该股票的 ETF，并展示 ETF 代码、市场、证券名称、替代标志等信息。
- 支持查看单只 ETF 的完整 PCF 成分。
- 支持下载单只 ETF 的 QMT 指定篮子 CSV。
- 支持一键将全市场已处理 PCF 成分打包为 CSV。
- 提供数据健康检查，按当次交易所 ETF 链接清单与有效 PCF 文件集合判断数据是否完整。

本地启动：

```powershell
python run_pcf_local.py
```

默认访问地址：

```text
http://127.0.0.1:5000/frontend/trade_tools.html
```

默认本地密码：

```text
local-test
```

更多本地启动和数据目录说明见 `docs/pcf-local.md`。

核心文件：

| 路径 | 说明 |
| --- | --- |
| `frontend/trade_tools.html` | ETF PCF 白名单前端页面 |
| `trade_tools/PCF/pcf_service.py` | PCF 链接抓取、文本下载、白名单匹配、成分解析和 CSV 导出服务 |
| `trade_tools/PCF/PCF_Whitelist_main.py` | PCF 白名单相关入口 |
| `run_pcf_local.py` | 本地启动器，自动选择端口并配置 PCF 数据目录 |
| `docs/pcf-local.md` | 本地启动和数据完整性说明 |

## 技术栈

- 后端：Python、Flask、Flask-CORS、Gunicorn
- 数据处理：pandas、openpyxl、xlrd、python-docx、pdfplumber、matplotlib
- 数据抓取：requests、Playwright
- 前端：原生 HTML/CSS/JavaScript 静态页面
- 部署：Render、Cloudflare Pages、GitHub Actions

## 后端接口

| 方法 | 路径 | 说明 | 鉴权 |
| --- | --- | --- | --- |
| `GET` | `/` | 健康检查 | 否 |
| `GET` | `/api/health` | 健康检查 | 否 |
| `POST` | `/api/invoice` | 发票分类 | 否 |
| `POST` | `/api/video` | 新股询价登记表生成 | 是 |
| `POST` | `/api/extra-revenue` | 超额收益回撤图生成 | 否 |
| `POST` | `/api/redemption` | 巨额赎回公告生成 | 是 |
| `POST` | `/api/valuation-table` | 估值表资产提取 | 是 |
| `POST` | `/api/pdf-withdraw/products` | 从 Excel 读取 PDF 抽取产品清单 | 是 |
| `POST` | `/api/pdf-withdraw` | 按产品清单抽取 PDF | 是 |
| `POST` | `/api/excel` | 关联交易决策留档生成 | 是 |
| `POST` | `/api/fund` | 关联交易公告生成 | 是 |
| `POST` | `/api/pcf/jobs` | 创建 PCF 白名单查询任务 | 是 |
| `GET` | `/api/pcf/jobs/<job_id>` | 查询 PCF 任务进度 | 是 |
| `GET` | `/api/pcf/jobs/<job_id>/result` | 获取 PCF 匹配结果 | 是 |
| `GET` | `/api/pcf/jobs/<job_id>/download` | 下载 PCF 匹配结果 Excel | 是 |
| `GET` | `/api/pcf/etfs/<market>/<etf_code>/components` | 查看 ETF PCF 成分 | 否 |
| `GET` | `/api/pcf/etfs/<market>/<etf_code>/qmt-basket` | 下载单只 ETF QMT 篮子 CSV | 否 |
| `GET` | `/api/pcf/etfs/all/components-csv` | 下载全市场 PCF 成分 CSV 包 | 是 |
| `POST` | `/api/pcf/crawl` | 同步创建 PCF 查询任务的兼容接口 | 是 |
| `GET` | `/api/pcf/status` | 查看 PCF 本地数据状态 | 否 |

鉴权接口支持两种 Token 传递方式：

```text
Authorization: Bearer <APP_SECRET_TOKEN>
X-App-Secret-Token: <APP_SECRET_TOKEN>
```

## 安全和文件处理

- 上传文件统一使用临时目录处理，处理完成后自动清理。
- ZIP 解压包含路径穿越保护，避免恶意压缩包写出目标目录。
- ZIP 文件名包含常见中文编码修复逻辑。
- 内部运营工具通过 `APP_SECRET_TOKEN` 保护。
- Flask CORS 显式暴露下载文件名和业务摘要响应头，便于静态前端读取。
- PCF 本地启动器默认只监听 `127.0.0.1`，本地数据目录默认放在项目外的持久目录。

## 本地运行

安装依赖：

```powershell
python -m pip install -r requirements.txt
```

运行 Flask 后端：

```powershell
$env:PORT="5000"
$env:APP_SECRET_TOKEN="local-test"
python app.py
```

访问页面：

```text
http://127.0.0.1:5000/frontend/IPO.html
http://127.0.0.1:5000/frontend/module_tools.html
http://127.0.0.1:5000/frontend/trade_tools.html
```

PCF 本地工具也可以直接通过启动器运行：

```powershell
python run_pcf_local.py
```

## 部署说明

后端部署到 Render 时建议使用：

```text
gunicorn app:app --bind 0.0.0.0:$PORT
```

需要配置的环境变量：

| 变量 | 说明 |
| --- | --- |
| `APP_SECRET_TOKEN` | 运营/交易内部接口鉴权 Token |
| `PORT` | Render 注入的监听端口 |
| `PCF_DATA_DIR` | 可选，PCF 持久化数据目录 |
| `PCF_USE_SYSTEM_PROXY` | 可选，是否让 PCF 抓取使用系统代理 |

前端静态文件可由 Cloudflare Pages 托管。GitHub Actions 会把 `frontend/` 和根目录跳转页复制到 `dist/` 后部署。

## 目录结构

```text
.
├─ app.py
├─ requirements.txt
├─ run_pcf_local.py
├─ .github/workflows/daily_crawl.yml
├─ crawlers/
├─ data/
│  └─ trading_holidays.json
├─ docs/
│  ├─ pcf-local.md
│  └─ agents/
├─ frontend/
│  ├─ IPO.html
│  ├─ module_tools.html
│  ├─ trade_tools.html
│  └─ unlock_db.js
├─ merge/
├─ module_tools/
├─ scripts/
├─ trade_tools/
└─ utils/
```

## 开发者协作说明

仓库保留 `AGENTS.md` 和 `docs/agents/`，用于记录 AI 编程助手的协作约定、Issue tracker 约定、triage label 词汇和领域文档布局。这些内容是开发协作配置，不是面向业务用户的功能模块。

个人安装的开源 Codex skill 不建议提交到本仓库。它们更适合保留在个人开发环境中，避免把第三方提示词、代理配置和非业务文件混入应用代码；如果未来确实要把某个 skill 作为项目标准工作流，应单独建文档说明它解决的开发问题、安装方式和适用场景。
