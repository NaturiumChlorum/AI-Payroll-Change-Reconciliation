# AI辅助薪酬服务变更核对系统

本地课堂演示 MVP：React + TypeScript + Vite / FastAPI / SQLite / pandas / openpyxl。
使用虚构或脱敏数据，完成“两份资料 → 字段映射 → 只读比较 → 规则解释 → 人工复核 → Excel 报告”闭环。
没有接入大模型，界面明确显示 **规则辅助（演示）**。无需 API Key，设置 LLM 密钥也不会启用模型调用。

## 本机直接启动（依赖及前端已安装构建）

在 PowerShell 中运行：

```powershell
cd 'C:\Users\Naturium Chlorum\Documents\Codex\2026-11-04\referenced-chatgpt-conversation-this-is-an\outputs\payroll-demo'
& '.\.venv312\Scripts\python.exe' -m uvicorn backend.main:app --host 127.0.0.1 --port 8000
```

打开 http://127.0.0.1:8000 。关闭终端或按 Ctrl+C 停止前台服务。若 8000 已占用，可以改成 8001，并打开对应地址。
也可运行 `powershell -ExecutionPolicy Bypass -File .\start.ps1`，或加 `-Port 8001`。
当前后台演示服务的日志位于 `logs/server.out.log` 和 `logs/server.err.log`，进程编号在 `logs/server.pid`。

## 首次安装 / 换电脑

要求：稳定版 Python 3.12、Node.js 22 或更新版本、npm。不要使用本机默认的 Python 3.14.0a5 测试版（其原生扩展会崩溃）。

本机明确选择已安装的 3.12：

```powershell
cd 'C:\Users\Naturium Chlorum\Documents\Codex\2026-11-04\referenced-chatgpt-conversation-this-is-an\outputs\payroll-demo'
& 'C:\Users\Naturium Chlorum\AppData\Local\Programs\Python\Python312\python.exe' -m venv .venv312
.\.venv312\Scripts\python.exe -m pip install -r requirements-lock.txt
cd frontend
npm.cmd ci
npm.cmd run build
cd ..
.\.venv312\Scripts\python.exe -m uvicorn backend.main:app --host 127.0.0.1 --port 8000
```

换电脑时将第一条 Python 路径替换为该电脑的 Python 3.12 路径，或用 `py -3.12 -m venv .venv312`。常规依赖范围在 requirements.txt，验证环境锁定版本在 requirements-lock.txt，前端使用 package-lock.json。

开发前端：保持后端 8000 运行，另开终端进入 frontend，执行 `npm.cmd run dev`，打开终端打印的本地地址（默认 5173）。Vite 将 /api 请求代理到 8000。单端口演示使用编译后的前端，无需运行 Vite。

## 五分钟演示

1. 点击“载入虚构样例”：变更前 6 行、变更后 8 行。也可以在“演示说明”下载两份 CSV，再分别通过文件选择框上传。`examples/` 同时提供 CSV 与 XLSX。
2. 主键设为“员工编号” ↔ “人员编码”。默认比较姓名、部门、月度金额和服务状态；每个旧字段可选择对应新字段或“不比较”。
3. 点击“开始只读核对”。预期：新增 1 项、缺失 1 项、重复 1 项、变更 2 项、空主键 1 项，未变更 2 个主键。
4. 查看 D002：月度金额 12000 → 15000，来源均为第 3 行。D003 有部门和服务状态两项变化。D005 为重复主键，展示全部来源记录，不任意选择一条配对。
5. 填写复核人、状态和备注（例如“课堂演示：已对照虚构变更通知，金额变化待确认”），保存。刷新页面后，在历史批次下拉框中重新打开，可验证保存结果。
6. 导出报告，打开 Excel 的“差异与复核”“批次说明”“操作日志”。到左侧“操作日志”查看导入、核对、人工复核、导出留痕。

样例均为“演示甲”等虚构姓名及 D 开头编号，与任何真实企业或员工无关。

## 比较规则和来源

- 主键去除首尾空格，区分大小写；空主键按源行单独列为异常，不互相匹配。
- 单侧或双侧出现重复主键时，整组列为重复，保留所有行；不同时再计新增、缺失或变更。
- 其他主键：仅新表出现为新增，仅旧表出现为缺失；双方唯一时比较所映射字段。
- 字段按原始文本精确比较，不自动转换金额、单位、日期或空格，因此 `100` 与 `100.00` 不同。CSV 保留前导零与 NA 文本；Excel 的数字显示格式不保留，带前导零的编号请存为文本。
- 一项差异对应一个主键；同一主键多个字段变化只计一项，报告则按变化字段展开。未变更仅针对所选字段。
- XLSX 只读第一张工作表；必须首行表头，不支持合并表头、重复/空表头、密码文件或公式。公式需先粘贴为值。不支持旧版 XLS。
- CSV 支持 UTF-8（含 BOM）和 GB18030，逗号分隔。来源行号是从表头为第 1 行开始的逻辑记录序号；含换行的引号单元格会占多条物理文本行。
- 每份最大 2 MB、5000 条记录、50 列，单元格最多 2000 字符。XLSX 解压后最大 20 MB。
- 报告把用户输入写成文本，不执行以等号开头的内容。输入文件不会被改写，解析后的记录保存在本地 SQLite；复核独立保存。

## 数据与环境变量

| 配置 | 默认值 | 说明 |
| --- | --- | --- |
| `DEMO_DB_PATH` | 项目目录下 `data/demo.sqlite3` | 可选 SQLite 文件路径；启动时读取。相对路径按启动目录解释，建议绝对路径。 |
| 主机 / 端口 | `127.0.0.1:8000` | 通过 uvicorn 命令参数设置，仅本机使用。 |

不自动读取 .env 文件。PowerShell 设置示例：`$env:DEMO_DB_PATH = 'C:\demo-data\classroom.sqlite3'`。
数据库存储导入记录、批次、复核意见和审计日志。网页显示最近 50 个批次、最近 200 条日志，数据库继续保留完整数据；导出日志最多包含该批次最近 200 条。

## 验证命令

```powershell
.\.venv312\Scripts\python.exe -m pytest -q
cd frontend
npm.cmd run build
cd ..
Invoke-RestMethod http://127.0.0.1:8000/api/health
```

自动测试覆盖样例计数、源行定位、跨字段映射、CSV/XLSX、前导零、NA、重复主键、异常文件、公式拒绝、复核持久化、导出文本安全和日志。测试使用独立临时数据库，不修改演示库。API 文档：http://127.0.0.1:8000/docs 。

## 范围与限制

这是单人本地课堂 MVP，无登录、权限管理、并发冲突控制、加密、数据库迁移和防篡改审计。复核人是手工填写的标签，不能证明真实身份。审计日志用于课堂追溯，不构成生产合规审计。多次复核以最后保存为准，历史修改留痕。
不实现工资计算、自动审批、源系统回写、支付、报税或外部消息。没有 OCR 或实时大模型。规则解释仅说明结构差异，不能判断薪酬的合法性或业务正确性。
本次仅交付网页及验证材料，尚未制作答辩 PPT，后续根据真实运行日志与截图调整。

## 目录

`backend/` 接口和核对逻辑；`frontend/` React 页面；`tests/` 测试；`examples/` 虚构样例；`data/` 本地数据库；`logs/` 运行日志；`docs/` 验证记录。
