# 🛒 拼多多订单导出工具
[![Python 3.8+](https://img.shields.io/badge/Python-3.8%2B-blue)]()
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)]()
将个人拼多多账号的**全部历史订单**抓取下来，自动分类整理，导出为带统计分析的 Excel 报表。支持按订单状态（全部 / 待付款 / 待分享 / 待发货 / 待收货 / 待评价）过滤抓取。
## ⚠️ 免责声明
- 本工具**仅供学习交流与个人数据备份**，请仅用于导出**自己的**订单数据
- 请勿用于商业用途或抓取他人数据
- 频繁请求可能触发平台风控，请自行评估风险（工具已内置请求间隔，请勿调小）
- 本项目与拼多多官方无关
## ✨ 功能特性
| 特性 | 说明 |
|---|---|
| 📄 全量抓取 | 基于订单接口游标分页，自动翻页直到抓完（首页 `offset` 传空串，之后逐页携带上一页末条订单的游标，返回空列表即结束） |
| 🖥️ 状态过滤 | 命令行 `-t/--type` 指定订单状态：全部 / 待付款 / 待分享 / 待发货 / 待收货 / 待评价，支持中文输入 |
| 🗂️ 自动分类 | 区分普通订单与合并单（多多买菜），生成"一单一行"与"一商品一行"两套视图 |
| 📊 Excel 报表 | 订单汇总 / 商品明细 / 统计分析 三个工作表 |
| 📈 统计分析 | 状态分布、月度消费趋势、店铺消费排行 Top30 |
| 💾 断点续传 | 每页实时落盘，中断或凭证失效后重跑自动去重续传，不同类型数据分文件互不干扰 |
| 🔒 凭证外置 | 敏感信息从环境变量 / `.env` 读取，内置零依赖加载器，不会泄露到仓库 |
| 🔁 失败重试 | 单页自动重试 + 递增退避 |
## 📦 安装
要求 Python 3.8+：
```bash
git clone https://github.com/<你的用户名>/pdd-orders-export.git
cd pdd-orders-export
pip install -r requirements.txt
```
## 🔑 配置凭证
需要 3 个凭证，均通过浏览器抓包获取（约 1 分钟）：
| 环境变量 | 抓包来源 |
|---|---|
| `PDD_UID` | 请求 URL 中的 `pdduid` 参数 |
| `PDD_ANTI_CONTENT` | 请求头 `anti-content` |
| `PDD_COOKIE` | 请求头 `cookie`（整行复制） |
**抓包步骤：**
1. Chrome / Edge 打开 [mobile.pinduoduo.com/orders.html](https://mobile.pinduoduo.com/orders.html) 并登录
2. 按 `F12` 打开开发者工具 → **Network（网络）** 面板
3. 刷新页面或向下滚动触发订单加载
4. 在请求列表中过滤找到 `order_list_v4` 请求
5. 按上表复制三个值
**方式一：`.env` 文件（推荐）**
```bash
cp .env.example .env    # Windows: copy .env.example .env
```
编辑 `.env` 填入三个值即可，程序启动时自动加载：
```env
PDD_UID=123456789
PDD_ANTI_CONTENT=0atyJC9Hxrz...
PDD_COOKIE=api_uid=...; pdd_user_id=...; ...
```
**方式二：系统环境变量**
```bash
# macOS / Linux
export PDD_UID="123456789"
export PDD_ANTI_CONTENT="0atyJC9Hxrz..."
export PDD_COOKIE="api_uid=...; ..."
# Windows PowerShell
$env:PDD_UID="123456789"
$env:PDD_ANTI_CONTENT="0atyJC9Hxrz..."
$env:PDD_COOKIE="api_uid=..."
```
## 🚀 运行
```bash
# 抓取全部订单（默认，等价于 -t all）
python pdd_orders.py
# 只抓指定状态的订单（-t / --type）
python pdd_orders.py -t unreceived      # 待收货
python pdd_orders.py --type 待付款       # 直接传中文也可以
python pdd_orders.py -t unpaidV2        # 接口原始值同样支持
# 查看帮助
python pdd_orders.py -h
```
**支持的订单状态**
| 命令行参数 | 含义 | 接口 type 值 |
|---|---|---|
| `all`（默认） | 全部 | `all` |
| `unpaid` | 待付款 | `unpaidV2` |
| `grouping` | 待分享 | `grouping` |
| `unshipping` | 待发货 | `unshipping` |
| `unreceived` | 待收货 | `unreceived` |
| `unrated` | 待评价 | `unrated` |
> 英文键、中文名、接口原始值均可，大小写不敏感；输错时会列出全部合法值。
**分类型存储说明**
不同订单类型的数据**分文件保存**，互不干扰：
| 命令 | 原始数据 | Excel 报表 |
|---|---|---|
| `-t all` | `pdd_orders_raw.json` | `拼多多订单报表.xlsx` |
| `-t unreceived` | `pdd_orders_raw_unreceived.json` | `拼多多订单报表_待收货.xlsx` |
| `-t unrated` | `pdd_orders_raw_unrated.json` | `拼多多订单报表_待评价.xlsx` |
断点续传也按类型独立：`-t unreceived` 中断后重跑同参数会自动续传，不会混入其他类型的数据。想强制重抓某类型，删除对应的 raw 文件即可。
控制台实时输出每单抓取进度，完成后在当前目录生成对应 Excel 报表。
## 📊 输出说明
**Sheet 1 · 订单汇总**（一单一行）
订单号 / 类型 / 店铺 / 商品概要 / 件数 / 实付 / 优惠 / 状态 / 下单·发货·收货时间 / 快递单号 / 备注
**Sheet 2 · 商品明细**（一商品一行）
订单号 / 合并单号 / 店铺 / 商品名称 / 规格 / 单价 / 数量 / 小计 / 商品ID / SKU / 状态 / 下单时间 / 图片链接（可点击）
**Sheet 3 · 统计分析**
消费总览（订单类型、总实付、总优惠、客单价、消费区间、数据完整性）· 按状态分布 · 按月消费趋势 · 店铺消费排行 Top30
## ⚙️ 可调参数（代码内常量）
| 参数 | 默认值 | 说明 |
|---|---|---|
| `PAGE_SIZE` | `10` | 每页条数 |
| `INTERVAL` | `1.5` | 请求间隔（秒），**不建议调小** |
| `MAX_RETRY` | `3` | 单页重试次数 |
| `RESUME` | `True` | 是否启用断点续传 |
订单类型请通过命令行 `-t/--type` 指定，见上文「运行」。
## ❓ 常见问题
**Q: 提示「未配置凭证」？**
按上文「配置凭证」一节，创建 `.env` 或设置环境变量。
**Q: 提示「返回非JSON / 登录态失效」？**
Cookie 过期。重新登录网页版并抓包，更新 `.env` 中的 `PDD_COOKIE`。
**Q: 提示「接口报错 error_code=...」？**
大概率是 `anti-content` 失效（该参数由前端 JS 生成、有时效性）。重新抓包更新 `.env`；已抓到的数据不会丢失，重跑自动续传。
**Q: 抓到一半中断了怎么办？**
直接重跑（同参数）。已抓订单保存在对应的 raw 文件中，程序自动去重续传。想强制重抓全量，删除该文件即可。
**Q: 切换到某个具体状态（如待评价）后接口报错？**
个别 tab 的请求与「全部」可能略有差异。请在网页上切到对应 tab 重新抓包，确认 `type` 取值与请求可复现；`anti-content` 本身也有时效性，失效就重新抓一份更新 `.env`。
**Q: 会被风控/封号吗？**
默认 1.5 秒/页的温和频率，正常完整导出一次通常没问题，但无法保证。请勿调小间隔或高频反复运行。
## 📁 项目结构
```
├── pdd_orders.py      # 主程序
├── .env.example       # 环境变量模板（复制为 .env 使用）
├── .gitignore
├── requirements.txt
└── README.md
```
运行后额外生成的文件（均已被 `.gitignore` 忽略）：
```
.env                          # 你的凭证（本地）
pdd_orders_raw*.json          # 各类型订单原始数据
拼多多订单报表*.xlsx           # 各类型 Excel 报表
```
## 📄 License
MIT