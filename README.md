# Procurement Automation Workbench

面向采购业务数据的 Windows 桌面自动化项目。项目把 Excel 整理、授权门户查询、附件下载与分类、分阶段回填、项目追踪表匹配和价格核对串成可恢复的工作流。

这是一个基于实习业务场景抽象出的个人项目作品。公开仓库**不包含**任何企业名称、真实网址、账号、登录会话、业务工作簿、供应商信息、页面扫描结果或运行日志。

### 公开演示数据边界

本公开版本仅使用固定的 `Demo Category` 与 `DEMO-MAT-*` 合成标识。测试中的项目、采购凭证和金额均为独立构造的演示值，不代表任何企业的物料、价格、成本、Saving 规则或业务结果。浏览器自动化仅在使用者自行提供已获授权的本地 URL 时才可运行；仓库测试不会登录门户、启动浏览器或访问业务文件。

## 项目解决的问题

采购数据处理往往横跨 Excel、企业门户和附件邮件：数据需要先整理成可分析的结构，再按采购凭证查询详情，下载并区分不同类型的附件，最后回填项目、价格和核对字段。本项目将这些步骤拆成独立阶段，支持在查询完成后再本地回填，也支持只重试缺失的采购凭证。

## 核心能力

| 模块 | 实现内容 |
| --- | --- |
| Excel 数据处理 | 读取采购记录、补充映射字段、生成数据表或透视表，并写入扩展业务列。 |
| 浏览器自动化 | 使用 Playwright 驱动 Edge，在已获授权的门户中按采购凭证查询详情。首次登录由使用者在浏览器中完成企业统一身份认证。 |
| 附件处理 | 下载详情页附件，按标准模板审批附件与常规附件分类；支持 `.msg`、`.eml`、PDF 等文件的后续解析路径。 |
| 可恢复流程 | 将“查询下载”和“本地解析回填”解耦，保存待回填清单；可仅补查缺失项，避免无差别重跑。 |
| 数据匹配与计算 | 以项目名称和内容字段匹配项目价格、成本与节省追踪表，回填价格、计划成本、节省额和差异核对结果。 |
| 桌面工作台 | Tkinter/ttkbootstrap 图形界面将各阶段操作放入同一工作台，并将耗时任务放到后台线程执行。 |

## 技术栈

- Python、pandas、openpyxl：Excel 读取、加工与回填
- pywin32：Windows Excel COM 自动化（原生透视表场景）
- Playwright：Edge 浏览器自动化和附件下载
- pdfplumber、extract-msg、Beautiful Soup：邮件与文档内容解析
- Tkinter、ttkbootstrap：桌面界面
- unittest：回归测试

## 工作流

1. 在工作台中选择本地采购记录，并生成目标数据表或透视表。
2. 补充项目、工作分解结构、附件及价格计算所需的扩展字段。
3. 在获得授权的业务门户中按采购凭证查询并下载附件。
4. 从待回填清单解析本地附件，回填可确认的字段；异常或歧义结果保留给人工核对。
5. 匹配项目价格、成本与节省追踪表，执行价格与差异计算。

查询阶段和回填阶段可以独立运行；若部分采购凭证未完成，可使用“补查缺失项”只处理缺口。

## 项目结构

```text
.
├── procurement_workbench.py # Excel 工作流与桌面工作台
├── web_query.py            # 门户查询、附件下载和回填逻辑
├── workbench_config.py     # 本地私有配置加载与校验
├── download_period.py      # 下载目录周期规则
├── run_workbench.py        # ASCII 名称的启动入口
├── 启动采购自动化.bat       # Windows 一键启动脚本
├── config.example.json     # 仅包含占位值的配置模板
└── tests/                  # 单元与回归测试
```

## 本地运行

### 前置条件

- Windows（图形界面、Edge 自动化和原生透视表场景依赖 Windows）
- Python 3.8+
- 如需生成原生 Excel 透视表，请在本机安装 Microsoft Excel
- 仅在你有权限访问的业务系统中使用浏览器自动化

### 配置私有门户地址

仓库故意不提供真实业务网址。复制模板后，在**本机**填写已获授权的地址；不要提交该文件。

```powershell
Copy-Item config.example.json config.local.json
```

`config.local.json` 中需要填写：

- `web.home_url`：业务门户首页
- `web.search_url`：采购查询页面

也可以使用环境变量 `PROCUREMENT_HOME_URL` 和 `PROCUREMENT_SEARCH_URL`。环境变量优先于本地配置文件。

### 启动

普通使用可双击 `启动采购自动化.bat`。脚本会在项目目录创建本机 `.venv`、安装依赖，并安装 Edge 自动化组件。

开发环境可手动执行：

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\playwright.exe install msedge
.\.venv\Scripts\python.exe run_workbench.py
```

业务工作簿、价格表、附件和登录状态均只应放在本机。仓库没有提供真实示例数据，因此请使用你有权访问且已脱敏的数据自行验证。

## 测试与验证

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -m py_compile procurement_workbench.py web_query.py workbench_config.py
```

公开版离线测试覆盖合成品类配置、`DEMO-MAT-*` 物料筛选和配置模板校验。它们使用临时文件和模拟 Excel 对象；不访问门户、不启动浏览器，也不等同于在真实业务环境中的端到端验收。

## 术语说明

| 公开表述 | 含义 |
| --- | --- |
| 企业统一身份认证 | 使用者在授权业务门户中完成的登录步骤。 |
| 审批请求编号 | 用于从查询结果进入对应详情页的业务标识。 |
| 项目管理系统项目名称 | 从详情页取得、用于后续项目追踪表匹配的名称字段。 |
| 标准模板审批附件 | 满足固定命名或格式规则、可自动解析的附件；不满足规则的附件会保留给人工处理。 |
| 项目价格、成本与节省追踪表 | 用于匹配项目、价格、计划成本与节省额的本地业务工作簿。 |

## 数据与安全边界

- `config.local.json`、登录状态、浏览器资料、下载附件、运行日志、页面快照和业务数据均由 `.gitignore` 排除。
- 不要提交账号、密码、Cookie、Token、私有 URL、供应商信息或包含真实采购数据的文件。
- 图像型 PDF 的 OCR 分支还依赖本机 Tesseract 和 Poppler；它们不是仓库的一部分。
- 公开版本不承诺业务效率指标、线上运行效果或真实系统截图；这些均需在授权环境内由使用者自行验证。

## 许可

当前仓库暂未附加开源许可证。若未来确认代码与素材的权属范围适合开放再使用，可另行选择合适的许可证。
