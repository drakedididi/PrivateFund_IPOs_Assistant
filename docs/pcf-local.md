# PCF 本地启动

在项目根目录执行：

```powershell
python run_pcf_local.py
```

启动成功后打开：

```text
http://127.0.0.1:5000/frontend/trade_tools.html
```

默认本地密码：

```text
local-test
```

启动器会自动查找项目外已有的 `pcf_runtime` 数据目录，只监听本机
`127.0.0.1`。如果 `5000` 端口已被占用，启动器会自动改用下一个空闲端口，
请以终端打印的“打开网页”地址为准。按 `Ctrl+C` 停止服务。

需要更换端口时：

```powershell
python run_pcf_local.py --port 5001
```

需要手动指定数据目录时：

```powershell
python run_pcf_local.py --data-dir "C:\path\to\pcf_runtime"
```

迁移到其他主机时，推荐始终使用 `--data-dir` 指向项目目录以外的持久目录。
如果没有指定且父级目录中也没有已有的 `pcf_runtime`，默认位置是：

- Windows：`%LOCALAPPDATA%\PrivateFundIPOsAssistant\pcf`
- macOS：`~/Library/Application Support/PrivateFundIPOsAssistant/pcf`
- Linux：`${XDG_DATA_HOME:-~/.local/share}/PrivateFundIPOsAssistant/pcf`

实际使用的完整路径会显示在网页的数据健康区域中。

## 数据完整性规则

系统不使用历史固定数量判断全市场 ETF 是否完整。每次刷新时，先从交易所
抓取当次 ETF 下载链接清单，再逐一下载 PCF。数据健康状态按当次链接清单与
有效 PCF 文件的 ETF 代码集合进行核对；只有代码逐一对应、文件可解析且日期
匹配时，才显示“数据完整”。因此新 ETF 上市后会自动进入当次检查范围。
