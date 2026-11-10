# -*- coding: utf-8 -*-
"""
拼多多 H5 订单列表 —— 分页抓取 + 数据分类 + Excel 报表导出
用法:
    python pdd_orders.py                  # 抓取全部订单(默认)
    python pdd_orders.py -t unreceived    # 只抓待收货
    python pdd_orders.py -t 待付款        # 传中文名同样支持
    python pdd_orders.py -h               # 查看帮助
分页逻辑: 首页 offset 传空串, 之后每次取上一页最后一条订单的 offset, 直到返回空列表
凭证: 从环境变量或 .env 文件读取 (PDD_UID / PDD_ANTI_CONTENT / PDD_COOKIE)
      模板见 .env.example, 抓包步骤见 README.md
"""
import argparse
import json
import os
import time
from collections import Counter, defaultdict
from datetime import datetime
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ═════════════════════ 0. 命令行参数 ═════════════════════
ORDER_TYPE_MAP = {
    # 命令行键       接口 type 值      中文名     额外别名
    "all": {"api": "all", "cn": "全部", "alias": []},
    "unpaid": {"api": "unpaidV2", "cn": "待付款", "alias": []},
    "grouping": {"api": "grouping", "cn": "待分享", "alias": []},
    "unshipping": {"api": "unshipping", "cn": "待发货", "alias": []},
    "unreceived": {"api": "unreceived", "cn": "待收货", "alias": []},
    "unrated": {"api": "unrated", "cn": "待评价", "alias": ["评价"]},
}


def resolve_order_type(value):
    """命令行输入 → 类型键。支持 英文键 / 接口原始值 / 中文名 / 别名，大小写不敏感"""
    v = str(value).strip().lower()
    for key, info in ORDER_TYPE_MAP.items():
        if v in (key, info["api"].lower(), info["cn"],
                 *(a.lower() for a in info["alias"])):
            return key
    return None


def parse_args():
    p = argparse.ArgumentParser(
        prog="pdd_orders.py",
        description="拼多多 H5 订单列表抓取 → 数据分类 → Excel 报表导出",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "订单状态类型:\n"
            "  all          全部订单 (默认)\n"
            "  unpaid       待付款\n"
            "  grouping     待分享\n"
            "  unshipping   待发货\n"
            "  unreceived   待收货\n"
            "  unrated      待评价\n"
            "  (也支持中文名或接口原始值，如: -t 待付款 / -t unpaidV2)\n"
            "\n"
            "示例:\n"
            "  python pdd_orders.py                  # 抓取全部订单\n"
            "  python pdd_orders.py -t unreceived    # 只抓待收货\n"
            "  python pdd_orders.py --type 待付款     # 中文同样有效\n"
            "\n"
            "凭证配置: 通过环境变量 PDD_UID / PDD_ANTI_CONTENT / PDD_COOKIE\n"
            "          或当前目录 .env 文件提供，详见 README.md"
        ),
    )
    p.add_argument("-t", "--type", default="all", metavar="TYPE",
                   help="订单状态过滤，见下方「订单状态类型」(默认: all)")
    args = p.parse_args()
    key = resolve_order_type(args.type)
    if key is None:
        valid = ", ".join(ORDER_TYPE_MAP)
        p.error(f"无效的订单类型 {args.type!r}，可选: {valid}（中文名亦可）")
    args.type_key = key
    return args


ARGS = parse_args()


# ═════════════════════ 1. 配置区 ═════════════════════
# ⚠️ 敏感凭证一律从环境变量 / .env 文件读取，禁止硬编码进代码！
# 变量名: PDD_UID / PDD_ANTI_CONTENT / PDD_COOKIE（模板见 .env.example）
def _load_dotenv(path=".env"):
    """极简 .env 加载器（零第三方依赖）。系统环境变量里已有的值优先，不会被覆盖。"""
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            key, val = key.strip(), val.strip()
            if len(val) >= 2 and val[0] == val[-1] and val[0] in ("'", '"'):
                val = val[1:-1]
            if key and key not in os.environ:
                os.environ[key] = val


_load_dotenv()
# ── 敏感参数（环境变量 / .env）──
PDDUID = os.getenv("PDD_UID", "").strip()
ANTI_CONTENT = os.getenv("PDD_ANTI_CONTENT", "").strip()
COOKIE_STR = os.getenv("PDD_COOKIE", "").strip()
# ── 抓取参数 ──
ORDER_TYPE = ORDER_TYPE_MAP[ARGS.type_key]["api"]  # 接口 type 值（由命令行 -t 决定）
ORDER_TYPE_CN = ORDER_TYPE_MAP[ARGS.type_key]["cn"]  # 中文类型名（用于显示和文件命名）
PAGE_SIZE = 10  # 每页条数
START_OFFSET = ""  # 起始游标，空=从最新开始抓全部
INTERVAL = 1.5  # 请求间隔(秒)，别调太小防风控
MAX_RETRY = 3  # 单页重试次数
# 不同订单类型的数据分文件保存，避免断点续传时互相混淆
_suffix = "" if ARGS.type_key == "all" else f"_{ARGS.type_key}"
RAW_FILE = f"pdd_orders_raw{_suffix}.json"
XLSX_FILE = "拼多多订单报表.xlsx" if ARGS.type_key == "all" else f"拼多多订单报表_{ORDER_TYPE_CN}.xlsx"
RESUME = True  # 中断后重跑自动续传
# ═══════════════════════════════════════════════════════════════════════════════
URL = f"https://mobile.pinduoduo.com/proxy/api/api/aristotle/order_list_v4?pdduid={PDDUID}"
HEADERS = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
    "anti-content": ANTI_CONTENT,
    "cache-control": "no-cache",
    "content-type": "application/json;charset=UTF-8",
    "cookie": COOKIE_STR,
    "origin": "https://mobile.pinduoduo.com",
    "pragma": "no-cache",
    "priority": "u=1, i",
    "referer": "https://mobile.pinduoduo.com/orders.html?type=0&comment_tab=1&combine_orders=1&main_orders=1&refer_page_name=personal&refer_page_id=10001_1794279462298_t895h409gf&refer_page_sn=10001",
    "sec-ch-ua": '"Not-A.Brand";v="99", "Microsoft Edge";v="157", "Chromium";v="157"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/157.0.0.0 Safari/537.36 Edg/157.0.0.0",
}
# 状态码兜底映射（优先使用接口返回的 order_status_prompt 文本）
STATUS_TEXT = {1: "待付款", 2: "待发货", 3: "待收货", 4: "已完成",
               5: "已完成", 6: "退款/售后", 7: "交易关闭", 8: "待评价"}


# ─────────────────────────── 工具函数 ───────────────────────────
def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def ts(t):
    """时间戳 → datetime（0/None 返回 None，Excel显示为空）"""
    return datetime.fromtimestamp(t) if t else None


def save_raw(orders):
    with open(RAW_FILE, "w", encoding="utf-8") as f:
        json.dump(orders, f, ensure_ascii=False)


def disp_w(v):
    """显示宽度: 中文按2算"""
    if v is None:
        return 0
    return sum(2 if ord(c) > 127 else 1 for c in str(v))


def autosize(ws, min_w=8, max_w=45):
    for col in ws.columns:
        best = max(disp_w(c.value) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(best + 2, min_w), max_w)


# ─────────────────────────── 抓取部分 ───────────────────────────
def build_payload(page, offset):
    return {
        "type": ORDER_TYPE,
        "page": page,
        "origin_host_name": "mobile.pinduoduo.com",
        "scene": "order_list_h5",
        "page_from": 0,
        "front_env": 1,
        "pay_front_supports": [],
        "anti_content": ANTI_CONTENT,
        "size": PAGE_SIZE,
        "offset": offset,
    }


def fetch_page(session, page, offset):
    body = json.dumps(build_payload(page, offset), ensure_ascii=False).encode("utf-8")
    resp = session.post(URL, headers=HEADERS, data=body, timeout=30)
    if resp.status_code != 200:
        raise RuntimeError(f"HTTP {resp.status_code}")
    if "json" not in resp.headers.get("content-type", ""):
        raise RuntimeError("返回非JSON(登录态/Cookie已失效)")
    data = resp.json()
    if isinstance(data, dict) and data.get("error_code") not in (None, 0):
        raise RuntimeError(f"接口报错 error_code={data.get('error_code')} {str(data.get('error_msg'))[:80]}")
    if not isinstance(data, dict):
        raise RuntimeError(f"响应格式异常: {str(data)[:100]}")
    return data


def order_brief(od):
    """控制台进度用的一行摘要"""
    if od.get("type") == 2:
        names = ";".join(g.get("goods_name", "")[:12]
                         for s in od.get("orders", []) for g in s.get("order_goods", []))
        return f"[合并单]{names[:28]}", (od.get("display_amount") or 0) / 100
    g0 = (od.get("order_goods") or [{}])[0]
    return g0.get("goods_name", "")[:30], (od.get("order_amount") or 0) / 100


def fetch_all():
    """返回 (订单列表, 是否完整抓完)"""
    print("═" * 60)
    print("  拼多多订单抓取 → Excel 报表")
    print(f"  账号UID : {PDDUID} | 类型: {ORDER_TYPE_CN}({ORDER_TYPE}) | 每页{PAGE_SIZE}条 | 间隔{INTERVAL}s")
    print(f"  输出    : {XLSX_FILE}")
    print("═" * 60)
    orders_all, seen = [], set()
    page, offset = 1, START_OFFSET
    # 断点续传
    if RESUME and os.path.exists(RAW_FILE):
        try:
            with open(RAW_FILE, encoding="utf-8") as f:
                orders_all = json.load(f)
            for od in orders_all:
                k = od.get("order_sn") or od.get("parent_order_sn")
                if k:
                    seen.add(k)
            if orders_all:
                offset = orders_all[-1].get("offset") or ""
                page = len(orders_all) // PAGE_SIZE + 1
                log(f"[续传] 本地已有 {len(orders_all)} 条，从第 {page} 页继续")
        except Exception:
            orders_all, seen = [], set()
    if not orders_all:
        log("从第 1 页开始抓取 ...")
    session = requests.Session()
    t0 = time.time()
    complete = False
    try:
        while True:
            log(f"── 第 {page} 页 | offset: {offset or '(空=从头开始)'}")
            data = None
            for attempt in range(1, MAX_RETRY + 1):
                try:
                    data = fetch_page(session, page, offset)
                    break
                except Exception as e:
                    log(f"   请求失败(第{attempt}/{MAX_RETRY}次): {e}")
                    if attempt < MAX_RETRY:
                        time.sleep(2 * attempt)
            if data is None:
                log("[X] 重试用尽 → 大概率 ANTI_CONTENT/Cookie 已失效，请重新抓包更新后重跑(支持续传)")
                break
            page_orders = data.get("orders") or []
            if not page_orders:
                log(f"[OK] 第 {page} 页返回空 → 已到最后一页，抓取完成")
                complete = True
                break
            # 去重入库 + 打印每条订单进度
            new_orders, page_amt = [], 0.0
            for od in page_orders:
                k = od.get("order_sn") or od.get("parent_order_sn")
                if k and k in seen:
                    continue
                if k:
                    seen.add(k)
                orders_all.append(od)
                new_orders.append(od)
                name, amt = order_brief(od)
                page_amt += amt
                log(f"    + {od.get('order_sn', '')}  ¥{amt:>8.2f}  {od.get('order_status_prompt', '')}  {name}")
            elapsed = time.time() - t0
            log(f"    本页 {len(page_orders)} 条 | 新增 {len(new_orders)} | 累计 {len(orders_all)} 单"
                f" | 本页金额 ¥{page_amt:.2f} | 已用时 {elapsed / 60:.1f} 分钟")
            if not new_orders:
                log("    本页全部重复，视为已到末尾")
                complete = True
                break
            next_offset = page_orders[-1].get("offset")
            if not next_offset:
                log("    本页未返回 offset，视为结束")
                complete = True
                break
            offset = next_offset
            save_raw(orders_all)  # 每页落盘，中途失效不丢数据
            if len(page_orders) < PAGE_SIZE:
                log(f"    提示: 本页 {len(page_orders)} 条 < 上限 {PAGE_SIZE} 条，下一页可能是最后一页")
            page += 1
            time.sleep(INTERVAL)
    except KeyboardInterrupt:
        print()
        log("[!] 手动中断，进度已保存（重跑自动续传）")
    save_raw(orders_all)
    return orders_all, complete


# ─────────────────────────── 数据分类提取 ───────────────────────────
def extract(orders):
    """
    分类规则:
      type=1 → 普通订单(单品/拼单)
      type=2 → 合并单(多多买菜等, 一个父单含多个子订单)
    输出: 订单汇总行(一单一行) + 商品明细行(一商品一行)
    """
    order_rows, goods_rows = [], []
    for od in orders:
        mall = (od.get("mall") or {}).get("mall_name", "") or "未知店铺"
        if od.get("type") == 2:
            # ── 合并单（多多买菜）──
            subs = od.get("orders") or []
            created = ts((od.get("group_order") or {}).get("create_at", 0))
            goods_flat = [g for s in subs for g in (s.get("order_goods") or [])]
            order_rows.append({
                "订单号": od.get("order_sn", ""),
                "类型": "多多买菜/合并单",
                "店铺": mall,
                "商品概要": " / ".join(g.get("goods_name", "") for g in goods_flat),
                "件数": sum(g.get("goods_number", 1) for g in goods_flat),
                "实付(元)": (od.get("display_amount") or 0) / 100,
                "优惠(元)": None,
                "状态": od.get("order_status_prompt", "") or "已下单",
                "下单时间": created, "发货时间": None, "收货时间": None,
                "快递单号": "",
                "备注": f"含{len(subs)}个子订单",
            })
            for s in subs:
                s_status = s.get("order_status_prompt") or od.get("order_status_prompt", "")
                for g in s.get("order_goods") or []:
                    p, q = g.get("goods_price") or 0, g.get("goods_number") or 1
                    goods_rows.append({
                        "订单号": s.get("order_sn", ""),
                        "合并单号": od.get("order_sn", ""),
                        "店铺": mall,
                        "商品名称": g.get("goods_name", ""),
                        "规格": g.get("spec", ""),
                        "单价(元)": p / 100, "数量": q, "小计(元)": p * q / 100,
                        "商品ID": g.get("goods_id", ""), "SKU": g.get("sku_id", ""),
                        "状态": s_status, "下单时间": created,
                        "图片链接": g.get("thumb_url", ""),
                    })
        else:
            # ── 普通订单 ──
            goods = od.get("order_goods") or []
            status = od.get("order_status_prompt") or STATUS_TEXT.get(od.get("status"), str(od.get("status", "")))
            group = od.get("group_order") or {}
            remark = "拼单" if group.get("customer_num", 1) > 1 else "单买"
            if od.get("after_sales"):
                remark += " | 有售后"
            order_rows.append({
                "订单号": od.get("order_sn", ""),
                "类型": "普通订单",
                "店铺": mall,
                "商品概要": " / ".join(g.get("goods_name", "") for g in goods),
                "件数": sum(g.get("goods_number", 1) for g in goods),
                "实付(元)": (od.get("order_amount") or 0) / 100,
                "优惠(元)": (od.get("discount_amount") or 0) / 100,
                "状态": status,
                "下单时间": ts(od.get("order_time", 0)),
                "发货时间": ts(od.get("shipping_time", 0)),
                "收货时间": ts(od.get("receive_time", 0)),
                "快递单号": od.get("tracking_number", ""),
                "备注": remark,
            })
            for g in goods:
                p, q = g.get("goods_price") or 0, g.get("goods_number") or 1
                goods_rows.append({
                    "订单号": od.get("order_sn", ""),
                    "合并单号": "",
                    "店铺": mall,
                    "商品名称": g.get("goods_name", ""),
                    "规格": g.get("spec", ""),
                    "单价(元)": p / 100, "数量": q, "小计(元)": p * q / 100,
                    "商品ID": g.get("goods_id", ""), "SKU": g.get("sku_id", ""),
                    "状态": status, "下单时间": ts(od.get("order_time", 0)),
                    "图片链接": g.get("thumb_url", ""),
                })
    return order_rows, goods_rows


# ─────────────────────────── Excel 导出 ───────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="E02E24")  # 拼多多红
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")


def write_table(ws, headers, rows, money_cols=(), dt_cols=()):
    money_cols, dt_cols = set(money_cols), set(dt_cols)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font, cell.fill, cell.alignment, cell.border = HEADER_FONT, HEADER_FILL, CENTER, BORDER
    for i, row in enumerate(rows, 2):
        for c, h in enumerate(headers, 1):
            v = row.get(h)
            if isinstance(v, str) and len(v) > 300:
                v = v[:300] + "..."
            cell = ws.cell(i, c, v)
            cell.border = BORDER
            if h in money_cols and isinstance(v, (int, float)):
                cell.number_format = "#,##0.00"
            elif h in dt_cols and isinstance(v, datetime):
                cell.number_format = "yyyy-mm-dd hh:mm"
            elif h == "图片链接" and isinstance(v, str) and v.startswith("http"):
                cell.hyperlink = v
                cell.font = Font(color="0563C1", underline="single")
    autosize(ws)
    ws.freeze_panes = "A2"  # 冻结表头


def write_stats(ws, order_rows, goods_rows, complete):
    r = [1]

    def sec(title):
        ws.cell(r[0], 1, title).font = Font(bold=True, size=12, color="E02E24")
        r[0] += 2

    def kv(k, v, money=False):
        ws.cell(r[0], 1, k).font = Font(bold=True)
        cell = ws.cell(r[0], 2, v)
        if money and isinstance(v, (int, float)):
            cell.number_format = "#,##0.00"
        r[0] += 1

    def table(headers, rows, money_cols=()):
        for c, h in enumerate(headers, 1):
            cell = ws.cell(r[0], c, h)
            cell.font, cell.fill, cell.alignment, cell.border = HEADER_FONT, HEADER_FILL, CENTER, BORDER
        r[0] += 1
        for row in rows:
            for c, v in enumerate(row, 1):
                cell = ws.cell(r[0], c, v)
                cell.border = BORDER
                if headers[c - 1] in money_cols and isinstance(v, (int, float)):
                    cell.number_format = "#,##0.00"
            r[0] += 1
        r[0] += 2

    n = len(order_rows)
    paid = sum(x["实付(元)"] or 0 for x in order_rows)
    disc = sum(x["优惠(元)"] or 0 for x in order_rows)
    dts = [x["下单时间"] for x in order_rows if x["下单时间"]]
    sec("一、消费总览")
    kv("报表生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    kv("订单类型", f"{ORDER_TYPE_CN}({ORDER_TYPE})")
    kv("数据完整性", "完整" if complete else "不完整(中途中断/凭证失效，可续传重跑补齐)")
    kv("订单总数", n)
    kv("  - 普通订单", sum(1 for x in order_rows if x["类型"] == "普通订单"))
    kv("  - 合并单(多多买菜等)", sum(1 for x in order_rows if x["类型"] != "普通订单"))
    kv("商品明细行数", len(goods_rows))
    kv("商品总件数", sum(x["件数"] or 0 for x in order_rows))
    kv("总实付(元)", round(paid, 2), money=True)
    kv("总优惠(元)", round(disc, 2), money=True)
    kv("平均客单价(元)", round(paid / n, 2) if n else 0, money=True)
    kv("首单时间", min(dts).strftime("%Y-%m-%d") if dts else "-")
    kv("最近下单", max(dts).strftime("%Y-%m-%d") if dts else "-")
    sec("二、按订单状态分布")
    st_cnt = Counter(x["状态"] or "未知" for x in order_rows)
    st_amt = defaultdict(float)
    for x in order_rows:
        st_amt[x["状态"] or "未知"] += x["实付(元)"] or 0
    table(["订单状态", "订单数", "实付合计(元)", "订单占比"],
          [[s, c, round(st_amt[s], 2), f"{c / n * 100:.1f}%"]
           for s, c in sorted(st_cnt.items(), key=lambda p: -p[1])],
          money_cols={"实付合计(元)"})
    sec("三、按月消费趋势")
    m_cnt, m_amt = Counter(), defaultdict(float)
    for x in order_rows:
        if x["下单时间"]:
            k = x["下单时间"].strftime("%Y-%m")
            m_cnt[k] += 1
            m_amt[k] += x["实付(元)"] or 0
    table(["月份", "订单数", "实付(元)"],
          [[k, m_cnt[k], round(m_amt[k], 2)] for k in sorted(m_cnt, reverse=True)],
          money_cols={"实付(元)"})
    sec("四、店铺消费排行 (Top 30)")
    p_cnt, p_amt = Counter(), defaultdict(float)
    for x in order_rows:
        k = x["店铺"] or "未知店铺"
        p_cnt[k] += 1
        p_amt[k] += x["实付(元)"] or 0
    top = sorted(p_cnt, key=lambda k: -p_amt[k])[:30]
    table(["排名", "店铺", "订单数", "实付(元)"],
          [[i + 1, k, p_cnt[k], round(p_amt[k], 2)] for i, k in enumerate(top)],
          money_cols={"实付(元)"})
    autosize(ws)


def write_excel(order_rows, goods_rows, complete):
    log("开始生成 Excel ...")
    wb = Workbook()
    log(f"  [1/3] 写入「订单汇总」 {len(order_rows)} 行 ...")
    ws1 = wb.active
    ws1.title = "订单汇总"
    write_table(ws1,
                ["订单号", "类型", "店铺", "商品概要", "件数", "实付(元)", "优惠(元)",
                 "状态", "下单时间", "发货时间", "收货时间", "快递单号", "备注"],
                order_rows,
                money_cols={"实付(元)", "优惠(元)"},
                dt_cols={"下单时间", "发货时间", "收货时间"})
    log(f"  [2/3] 写入「商品明细」 {len(goods_rows)} 行 ...")
    ws2 = wb.create_sheet("商品明细")
    write_table(ws2,
                ["订单号", "合并单号", "店铺", "商品名称", "规格", "单价(元)", "数量",
                 "小计(元)", "商品ID", "SKU", "状态", "下单时间", "图片链接"],
                goods_rows,
                money_cols={"单价(元)", "小计(元)"},
                dt_cols={"下单时间"})
    log("  [3/3] 写入「统计分析」 ...")
    ws3 = wb.create_sheet("统计分析")
    write_stats(ws3, order_rows, goods_rows, complete)
    wb.save(XLSX_FILE)
    log(f"Excel 已保存: {XLSX_FILE}")


def print_summary(order_rows, goods_rows, complete):
    n = len(order_rows)
    paid = sum(x["实付(元)"] or 0 for x in order_rows)
    disc = sum(x["优惠(元)"] or 0 for x in order_rows)
    dts = [x["下单时间"] for x in order_rows if x["下单时间"]]
    p_amt = defaultdict(float)
    for x in order_rows:
        p_amt[x["店铺"] or "未知"] += x["实付(元)"] or 0
    top_mall, top_amt = max(p_amt.items(), key=lambda p: p[1]) if p_amt else ("-", 0)
    line = "═" * 60
    print()
    print(line)
    print("  抓 取 结 果 摘 要")
    print(line)
    print(f"  订单类型   : {ORDER_TYPE_CN}")
    print(f"  数据完整性 : {'完整' if complete else '不完整(重跑可续传补齐)'}")
    print(f"  订单总数   : {n} 单 (普通 {sum(1 for x in order_rows if x['类型'] == '普通订单')}"
          f" + 合并单 {sum(1 for x in order_rows if x['类型'] != '普通订单')})")
    print(f"  商品明细   : {len(goods_rows)} 行")
    print(f"  总实付     : ¥{paid:,.2f}    总优惠: ¥{disc:,.2f}")
    if dts:
        print(f"  消费区间   : {min(dts):%Y-%m-%d} ~ {max(dts):%Y-%m-%d}")
    print(f"  消费最多   : {top_mall} (¥{top_amt:,.2f})")
    print(f"  输出文件   : {XLSX_FILE}")
    print(f"  原始数据   : {RAW_FILE}")
    print(line)


# ─────────────────────────── 凭证校验 ───────────────────────────
def check_credentials():
    """启动前校验凭证是否已配置，未配置则给出指引并退出"""
    missing = [name for name, val in (
        ("PDD_UID", PDDUID),
        ("PDD_ANTI_CONTENT", ANTI_CONTENT),
        ("PDD_COOKIE", COOKIE_STR),
    ) if not val]
    if not missing:
        return True
    print("═" * 60)
    print("  [X] 未配置凭证，缺少环境变量: " + ", ".join(missing))
    print()
    print("  配置方法（二选一）:")
    print("   1. 复制 .env.example 为 .env，填入抓包得到的三个值（推荐，程序自动加载）")
    print("   2. 直接设置系统环境变量:")
    print('      macOS/Linux : export PDD_UID="..." PDD_ANTI_CONTENT="..." PDD_COOKIE="..."')
    print('      PowerShell  : $env:PDD_UID="..."; $env:PDD_ANTI_CONTENT="..."; $env:PDD_COOKIE="..."')
    print()
    print("  凭证获取步骤见 README.md「配置凭证」一节")
    print("═" * 60)
    return False


# ─────────────────────────── 主流程 ───────────────────────────
def main():
    if not check_credentials():
        return
    orders, complete = fetch_all()
    if not orders:
        log("未获取到任何订单数据，退出")
        return
    log(f"开始分类整理 {len(orders)} 个订单 ...")
    order_rows, goods_rows = extract(orders)
    # 按下单时间倒序
    order_rows.sort(key=lambda x: x["下单时间"] or datetime.min, reverse=True)
    goods_rows.sort(key=lambda x: x["下单时间"] or datetime.min, reverse=True)
    log(f"分类完成: 订单汇总 {len(order_rows)} 条 | 商品明细 {len(goods_rows)} 行")
    write_excel(order_rows, goods_rows, complete)
    print_summary(order_rows, goods_rows, complete)


if __name__ == "__main__":
    main()