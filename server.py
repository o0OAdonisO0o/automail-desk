# -*- coding: utf-8 -*-
r"""Automail Desk —— 外链采购发信 / 回信处理面板（后端，单文件标准库实现）

目录结构（均相对本文件所在目录，可用 AUTOMAIL_BASE 覆盖）
- data/automail.db   状态库（账号 / 队列 / 发信 / 回信 / 链接检查）
- ui/index.html      面板
- exports/           A 类客户文档、每日日志
- logs/              运行日志
- config.json        配置（域名、密码、邮件服务商、账号池）

启动：python server.py          # 服务 + 调度 + 收信轮询
      python server.py seed     # 仅初始化数据（账号池 + 收件人池）
"""
import csv, imaplib, json, os, random, re, socket, sqlite3, ssl, sys, threading, time, traceback
import smtplib
import urllib.request, urllib.error
from datetime import datetime, date, timedelta
from email import message_from_bytes, policy
from email.utils import parseaddr, getaddresses
from email.message import EmailMessage
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, unquote

BASE = os.path.abspath(os.environ.get("AUTOMAIL_BASE") or os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE, "data", "automail.db")
UI = os.path.join(BASE, "ui", "index.html")
EXPORT_DIR = os.path.join(BASE, "exports")
LOG_PATH = os.path.join(BASE, "logs", "server.log")


def load_config():
    """读取 BASE/config.json；环境变量优先级更高，缺省值面向首次运行。"""
    path = os.path.join(BASE, "config.json")
    cfg = {}
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                cfg = json.load(f) or {}
        except Exception as e:
            print(f"[config] config.json 解析失败（按默认值运行）：{e}", flush=True)
    env = os.environ
    for key, name in (("domain", "AUTOMAIL_DOMAIN"), ("password", "AUTOMAIL_PASSWORD"),
                      ("smtp_host", "AUTOMAIL_SMTP_HOST"), ("imap_host", "AUTOMAIL_IMAP_HOST")):
        if env.get(name):
            cfg[key] = env[name]
    if env.get("AUTOMAIL_PORT"):
        cfg["port"] = env["AUTOMAIL_PORT"]
    return cfg


CFG = load_config()

DOMAIN = (CFG.get("domain") or "example.com").strip()
PASSWORD = CFG.get("password") or ""
SMTP_HOST = CFG.get("smtp_host") or "smtp.qiye.aliyun.com"
SMTP_PORT = int(CFG.get("smtp_port") or 465)
IMAP_HOST = CFG.get("imap_host") or "imap.qiye.aliyun.com"
IMAP_PORT = int(CFG.get("imap_port") or 993)
XLSX = CFG.get("recipient_xlsx") or ""
ACCOUNT_PREFIX = CFG.get("account_prefix") or "send"
ACCOUNT_COUNT = int(CFG.get("account_count") or 0)
PORT = int(os.environ.get("AUTOMAIL_PORT") or CFG.get("port") or 8765)
READONLY = os.environ.get("AUTOMAIL_READONLY") == "1"   # 沙箱/自测：不启动调度与收信线程

_lock = threading.Lock()


def log(*a):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] " + " ".join(str(x) for x in a)
    print(line, flush=True)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def conn():
    c = sqlite3.connect(DB_PATH, timeout=30)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA journal_mode=WAL")
    return c


SCHEMA = """
CREATE TABLE IF NOT EXISTS accounts(
  id INTEGER PRIMARY KEY, email TEXT UNIQUE, pwd TEXT,
  status TEXT DEFAULT 'ok', sent_total INTEGER DEFAULT 0, sent_today INTEGER DEFAULT 0,
  last_send_at TEXT, reset_date TEXT, last_error TEXT);
CREATE TABLE IF NOT EXISTS recipients(
  id INTEGER PRIMARY KEY, email TEXT UNIQUE, name TEXT, domain TEXT, company TEXT,
  source_id TEXT, status TEXT DEFAULT 'new', category TEXT DEFAULT '',
  round_done INTEGER DEFAULT 0, last_sent_at TEXT, account_id INTEGER,
  reply_at TEXT, note TEXT, auto INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS messages(
  id INTEGER PRIMARY KEY, account_id INTEGER, recipient_id INTEGER, rnd INTEGER,
  tpl TEXT, subject TEXT, message_id TEXT, status TEXT, sent_at TEXT, error TEXT);
CREATE TABLE IF NOT EXISTS replies(
  id INTEGER PRIMARY KEY, account_id INTEGER, message_id TEXT, from_addr TEXT,
  subject TEXT, received_at TEXT, body TEXT, category TEXT, auto INTEGER DEFAULT 0,
  is_bounce INTEGER DEFAULT 0, links_json TEXT, quote_hit INTEGER DEFAULT 0, snippet TEXT);
CREATE TABLE IF NOT EXISTS link_checks(
  id INTEGER PRIMARY KEY, reply_id INTEGER, url TEXT, status_code INTEGER,
  final_url TEXT, flags TEXT, checked_at TEXT);
CREATE TABLE IF NOT EXISTS folders(account_id INTEGER, folder TEXT, last_uid INTEGER, PRIMARY KEY(account_id,folder));
CREATE TABLE IF NOT EXISTS domain_pool(
  id INTEGER PRIMARY KEY, domain TEXT UNIQUE, source TEXT, added_at TEXT,
  expanded INTEGER DEFAULT 0, addr_added INTEGER DEFAULT 0, note TEXT);
CREATE TABLE IF NOT EXISTS settings(k TEXT PRIMARY KEY, v TEXT);
CREATE TABLE IF NOT EXISTS events(id INTEGER PRIMARY KEY, ts TEXT, kind TEXT, detail TEXT);
"""

DEFAULT_TEMPLATES = {
    "S1": {
        "subject": "Publishing partnership with {domain}",
        "body": """Hi {first_name},

We are expanding our publisher network and would like to establish a long-term purchasing partnership with {domain} for recurring client projects.

Could you please share your current publishing rates and conditions?

- Pricing for guest posts, link insertions, sponsored posts and content writing (if available).
- Accepted niches and separate rates for general content, casino/iGaming, sports betting, crypto, forex, CBD, adult and other restricted categories.
- Current DA, DR, monthly organic traffic and main traffic countries, preferably with the data source.
- Link type (dofollow/nofollow), number of links allowed, placement duration and sponsored disclosure requirements.
- Current availability, turnaround time and any editorial restrictions.
- Payment methods, post-publication payment or NET30 options, and best rates for recurring orders.

We are looking for a reliable partner who can accept orders on an ongoing basis. Could you confirm how long your quoted rates are valid and whether availability needs to be checked before each order?

Please notify us in advance of any changes to pricing, availability or publishing conditions so we can keep our records accurate.

We look forward to building a long-term partnership.

{signature}""",
    },
    "S2": {
        "subject": "Publishing rates and availability - {domain}",
        "body": """Hi {first_name},

We are interested in purchasing recurring placements on {domain} and would like to establish a long-term partnership.

Could you share your current guest post, link insertion, sponsored post and content writing rates, along with DA, DR, monthly organic traffic and main traffic countries?

Please also confirm which niches you accept (general, casino/iGaming, sports betting, crypto, forex, CBD, adult, etc.) and any additional charges or restrictions.

We would appreciate details on current availability, link conditions, placement duration, turnaround time, payment options and your best recurring-order rates. We prefer payment after publication and are open to NET30 arrangements.

How long are your rates valid? Please let us know in advance if pricing, availability or publishing rules change.

{signature}""",
    },
    "S3": {
        "subject": "Long-term media sourcing partnership - {domain}",
        "body": """Hi {first_name},

We are expanding our media supplier network and are interested in sourcing placements on {domain}, as well as other websites you can reliably provide.

Could you share your latest website inventory in Excel or Google Sheets, including the following details for each site?

- Website URL, DA, DR, monthly organic traffic, main traffic countries and metric sources.
- Separate reseller rates for guest posts, link insertions, sponsored posts and content writing.
- Accepted niches and separate prices for general, casino/iGaming, sports betting, crypto, forex, CBD, adult and other restricted categories.
- Link attributes, number of links allowed, placement duration, disclosure requirements and editorial restrictions.
- Current order availability, turnaround time and whether you work directly with publishers or through other suppliers.
- Payment methods, post-publication payment or NET30 options, and volume discounts.

Our priority is reliable delivery and competitive, consistent pricing for recurring orders. Please confirm the validity of your rates and whether individual orders require an availability check.

We would appreciate advance notice of any price changes, unavailable websites or changes in publishing requirements.

We are looking for long-term cooperation rather than one-off orders.

{signature}""",
    },
    "F1": {
        "subject": "Re: {original_subject}",
        "body": """Hi {first_name},

Just following up on my earlier inquiry about publishing placements on {domain}.

We are still interested in a long-term partnership. If you have a current rate card or website list covering publishing options, accepted niches and availability, please feel free to share it.

We would also appreciate confirmation of price validity, payment options and how you notify partners of changes to rates or availability.

If you are not the right contact, could you kindly direct me to the person responsible for publishing partnerships?

{signature}""",
    },
    "F2": {
        "subject": "Re: {original_subject}",
        "body": """Hi {first_name},

Thank you for sharing your rates. Before adding {domain} to our purchasing inventory, could you confirm that the quoted services are currently available for orders, subject to your editorial requirements?

Please confirm the price validity period, any niche-specific charges, whether each order requires separate availability confirmation, and the expected turnaround time.

Could you also confirm your payment terms and whether you can notify us before any changes to pricing, availability or publishing rules?

If available, please share a few recent live examples of the relevant placement types.

We look forward to working together on recurring orders.

{signature}""",
    },
}

DEFAULT_SETTINGS = {
    "campaign_state": "idle",
    "cooldown_min": "15",
    "daily_cap": "3",
    "followup_cap": "2",
    "followup_gap_hours": "24",
    "rounds": "2",
    "poll_sec": "180",
    # 发送窗口与节奏
    "send_window_start": "09:00",
    "send_window_end": "20:00",
    "send_gap_sec": "20",
    "max_per_domain_day": "3",
    # 风控与自动化
    "bounce_pause_pct": "15",
    "bounce_min_sample": "30",
    "scan_spam": "1",
    "link_check": "1",
    "ai_auto_queue": "1",
    # 域名展开
    "domain_local_parts": "info,contact,sales",
    "expand_on_upload": "1",
    "sender_name": CFG.get("sender_name") or "Leo Li",
    "company_name": CFG.get("company_name") or "GlobalSMS Connect",
    "whatsapp": "",
    "optout": "If you would prefer not to receive further messages, reply with \"unsubscribe\" and we will remove your address from our list.",
    "templates": json.dumps(DEFAULT_TEMPLATES, ensure_ascii=False),
}


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    with conn() as c:
        c.executescript(SCHEMA)
        # 旧库补列：早期版本建表时没有 AI 判定字段
        cols = {r[1] for r in c.execute("PRAGMA table_info(replies)")}
        for name, ddl in (("ai_done", "INTEGER DEFAULT 0"), ("ai_json", "TEXT")):
            if name not in cols:
                c.execute(f"ALTER TABLE replies ADD COLUMN {name} {ddl}")
                log("migrate: replies." + name)
        for k, v in DEFAULT_SETTINGS.items():
            c.execute("INSERT OR IGNORE INTO settings(k,v) VALUES(?,?)", (k, v))
    log("DB ready:", DB_PATH)


def get_settings():
    with conn() as c:
        d = {r["k"]: r["v"] for r in c.execute("SELECT k,v FROM settings")}
    d["templates"] = json.loads(d.get("templates") or "{}")
    return d


def set_settings(patch: dict):
    with conn() as c:
        for k, v in patch.items():
            if k == "templates" and not isinstance(v, str):
                v = json.dumps(v, ensure_ascii=False)
            c.execute("INSERT INTO settings(k,v) VALUES(?,?) ON CONFLICT(k) DO UPDATE SET v=excluded.v",
                      (k, str(v)))


# ---------------------------------------------------------------- 智能闸门
def send_window_open(s=None):
    """是否处于允许发信的时段（跨零点区间亦支持）"""
    s = s or get_settings()
    try:
        a = s.get("send_window_start") or "00:00"
        b = s.get("send_window_end") or "23:59"
        now = datetime.now().strftime("%H:%M")
        if a == b:
            return True
        return (a <= now <= b) if a < b else (now >= a or now <= b)
    except Exception:
        return True


def today_stats():
    today = date.today().isoformat()
    sent = q("SELECT COUNT(*) c FROM messages WHERE sent_at LIKE ? AND status='sent'", (today + "%",))[0]["c"]
    bad = q("SELECT COUNT(*) c FROM messages WHERE sent_at LIKE ? AND status!='sent'", (today + "%",))[0]["c"]
    total = sent + bad
    return {"sent": sent, "bad": bad, "total": total, "rate": round(bad * 100.0 / total, 1) if total else 0.0}


def bounce_guard(s=None):
    """退信率超过阈值则自动暂停。返回 (要不要拦, 说明)"""
    s = s or get_settings()
    try:
        limit = float(s.get("bounce_pause_pct") or 0)
        min_n = int(s.get("bounce_min_sample") or 0)
    except Exception:
        return False, ""
    if limit <= 0:
        return False, ""
    st = today_stats()
    if st["total"] < max(1, min_n):
        return False, ""
    if st["rate"] >= limit:
        set_settings({"campaign_state": "paused"})
        event("auto_pause", f"退信率 {st['rate']}% 达到阈值 {limit}%，已自动暂停")
        log(f"AUTO-PAUSE 退信率 {st['rate']}% >= {limit}%")
        return True, f"退信率 {st['rate']}%（{st['bad']}/{st['total']}）达到阈值 {limit}%，已自动暂停"
    return False, ""


def preflight():
    """发信前体检：把会导致整批失败或进垃圾箱的问题提前暴露"""
    s = get_settings()
    items = []

    def add(level, title, detail):
        items.append({"level": level, "title": title, "detail": detail})

    n_acc = q("SELECT COUNT(*) c FROM accounts WHERE status='ok'")[0]["c"]
    n_rec = q("SELECT COUNT(*) c FROM recipients WHERE status='new' AND round_done=0")[0]["c"]
    cap = int(s.get("daily_cap") or 0) * max(1, n_acc)
    if n_acc == 0:
        add("bad", "无可用发信账号", "先在「收件人与域名」页上传账号池")
    else:
        add("ok", "发信账号", f"{n_acc} 个可用")
    if n_rec == 0:
        add("warn", "无待发收件人", "上传邮箱名单或域名名单后才会进入调度")
    else:
        add("ok", "待发收件人", f"{n_rec} 个；按当前配额今日可承载 {cap} 封")
        if n_rec < cap:
            add("warn", "收件人不足以吃满配额", f"待发 {n_rec} < 今日容量 {cap}，会提前停摆")

    st = today_stats()
    limit = float(s.get("bounce_pause_pct") or 0)
    if st["total"] >= 10 and st["rate"] >= limit > 0:
        add("bad", "退信率超阈值", f"今日 {st['rate']}%（{st['bad']}/{st['total']}），调度器已暂停")
    elif st["total"]:
        add("ok" if st["rate"] < 10 else "warn", "退信率", f"今日 {st['rate']}%（{st['bad']}/{st['total']}）")

    # 模板占位符残留
    sample = {"email": "probe@example.com", "name": "Probe", "domain": "example.com", "company": "Probe Co"}
    residue = []
    for k in (s.get("templates") or {}):
        try:
            subj, body = render(k, sample, s, "old subject")
            if "{" in subj + body or "}" in subj + body:
                residue.append(k)
        except Exception:
            residue.append(k + "(渲染失败)")
    if residue:
        add("bad", "模板存在未替换占位符", "、".join(residue))
    else:
        add("ok", "模板渲染", f"{len(s.get('templates') or {})} 套模板占位符全部可替换")

    if not (s.get("whatsapp") or "").strip():
        add("warn", "WhatsApp 未配置", "签名中不输出该行；如需客户直接联系请在配置页填写")
    if s.get("scan_spam") != "1":
        add("warn", "垃圾箱未纳入收信", "部分回信可能落在垃圾箱而漏收")
    if s.get("link_check") != "1":
        add("warn", "回信链接未检查", "无法识别可疑跳转与短链")

    # 域名集中度
    rows = q("""SELECT domain, COUNT(*) c FROM recipients GROUP BY domain ORDER BY c DESC LIMIT 1""")
    if rows and n_rec:
        top = rows[0]
        if top["c"] >= max(5, n_rec * 0.4):
            add("warn", "收件域名过于集中", f"{top['domain']} 占 {top['c']}/{n_rec}，易被按域拦截")
    if s.get("campaign_state") == "running" and not send_window_open(s):
        add("warn", "当前在发送时段之外", f"窗口 {s.get('send_window_start')}–{s.get('send_window_end')}，调度器已挂起")
    n_ai = q("SELECT COUNT(*) c FROM replies WHERE IFNULL(ai_done,0)=0 AND is_bounce=0")[0]["c"]
    if n_ai:
        add("warn", "待识别回信", f"{n_ai} 封等待 WorkBuddy 判定 A/B/C")
    bad = sum(1 for i in items if i["level"] == "bad")
    return {"score": max(0, 100 - bad * 30 - sum(1 for i in items if i["level"] == "warn") * 5),
            "items": items, "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}


def event(kind, detail):
    with conn() as c:
        c.execute("INSERT INTO events(ts,kind,detail) VALUES(?,?,?)",
                  (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), kind, detail))


# ---------------------------------------------------------------- seed
def seed():
    """初始化数据：按 config.json 的 account_prefix/account_count 生成账号池，
    若配置了 recipient_xlsx 则同时导入收件人池。"""
    init_db()
    with conn() as c:
        for i in range(ACCOUNT_COUNT):
            acct = f"{ACCOUNT_PREFIX}{i:03d}@{DOMAIN}"
            c.execute("INSERT OR IGNORE INTO accounts(email,pwd,reset_date) VALUES(?,?,?)",
                      (acct, PASSWORD, date.today().isoformat()))
        n = 0
        if XLSX and os.path.exists(XLSX):
            try:
                import openpyxl
            except ImportError:
                log("未安装 openpyxl，跳过收件人表导入；可执行 pip install -r requirements.txt")
                openpyxl = None
            if openpyxl:
                wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
                ws = wb[wb.sheetnames[0]]
                for r in ws.iter_rows(min_row=2, values_only=True):
                    email = str(r[1] or "").strip().lower() if len(r) > 1 else ""
                    if "@" not in email:
                        continue
                    dom = str(r[5] or "").strip().lower() if len(r) > 5 else ""
                    dom = dom or email.split("@")[1]
                    cur = c.execute(
                        "INSERT OR IGNORE INTO recipients(email,name,domain,company,source_id) VALUES(?,?,?,?,?)",
                        (email, str(r[2] or "").strip() if len(r) > 2 else "", dom,
                         str(r[4] or "").strip() if len(r) > 4 else "",
                         str(r[11] or "") if len(r) > 11 else ""))
                    n += cur.rowcount
        elif XLSX:
            log("recipient_xlsx 不存在，跳过导入：", XLSX)
        acc = c.execute("SELECT COUNT(*) FROM accounts").fetchone()[0]
        rec = c.execute("SELECT COUNT(*) FROM recipients").fetchone()[0]
    log(f"seed done: accounts={acc} recipients={rec} (new {n})")
    print(f"账号 {acc} 个，收件人 {rec} 个")


# ---------------------------------------------------------------- 模板渲染
def signature(s):
    parts = [f"Best regards,\n{s['sender_name']}", s["company_name"]]
    if s.get("whatsapp"):
        parts.append(f"WhatsApp: {s['whatsapp']}")
    if s.get("optout"):
        parts.append(s["optout"])
    return "\n".join(parts)


def pick_tpl(rec, rnd, s):
    if rnd == 1:
        local = rec["email"].split("@")[0].lower()
        generic = local in {"info", "iletisim", "contact", "admin", "sales", "support",
                            "hello", "office", "mail", "enquiry", "inquiry", "press"}
        if rec["name"] and not generic:
            return "S1"
        return "S2"
    # 二轮起：有过回复用 F2，未回复用 F1
    return "F2" if (rec["reply_at"] and rec["category"] in ("A", "B")) else "F1"


def render(tpl_key, rec, s, original_subject=""):
    t = s["templates"].get(tpl_key) or DEFAULT_TEMPLATES[tpl_key]
    first = (rec["name"] or "").strip() or "there"
    # 只取名字首段，避免 "Ahmet Yilmaz" 之类整段塞进 Hi
    first = first.split()[0] if first != "there" else first
    mapping = {
        "domain": rec["domain"],
        "first_name": first,
        "company": rec["company"] or rec["domain"],
        "whatsapp": s.get("whatsapp", ""),
        "signature": signature(s),
        "original_subject": original_subject or f"Publishing partnership with {rec['domain']}",
    }
    def sub(txt):
        for k, v in mapping.items():
            txt = txt.replace("{" + k + "}", v)
        return txt
    return sub(t["subject"]), sub(t["body"])


# ---------------------------------------------------------------- 发信
def smtp_send(account_email, pwd, to_addr, subject, body, msgid):
    msg = EmailMessage()
    msg["From"] = f"{account_email}"
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg["Message-ID"] = msgid
    msg["Date"] = datetime.now().astimezone().strftime("%a, %d %b %Y %H:%M:%S %z")
    msg.set_content(body)
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30, context=ctx) as s:
        s.login(account_email, pwd)
        s.send_message(msg)


def due_recipients(c, rnd, exclude_domain_by_account=None, account_id=None):
    """取该轮次可发收件人：已完成 rnd-1 轮、未退信、非 A 类（rnd>1）"""
    sql = """SELECT * FROM recipients
             WHERE status NOT IN ('bounced','unsubscribed','done')
               AND round_done = ?"""
    args = [rnd - 1]
    if rnd > 1:
        sql += " AND (category IS NULL OR category != 'A')"
        # 跟进须与上一封间隔足够久，避免刚发完就跟
        gap = int(get_settings().get("followup_gap_hours", "24"))
        cutoff = (datetime.now() - timedelta(hours=gap)).strftime("%Y-%m-%d %H:%M:%S")
        sql += " AND IFNULL(last_sent_at,'') <= ?"
        args.append(cutoff)
    if account_id is not None:
        sql += " AND (account_id IS NULL OR account_id != ?)"
        args.append(account_id)
    rows = c.execute(sql + " ORDER BY (name!='') DESC, id", args).fetchall()
    return rows


def send_one(job):
    """job=(account_row, recipient_row, rnd)

    先写 sending 再发信：进程若在 SMTP 往返中崩溃，重启后该行由 recover_inflight()
    标为 unknown 并跳过重发，宁可漏发也不重复打扰收件人。
    """
    s = get_settings()
    with conn() as c:
        acc = c.execute("SELECT * FROM accounts WHERE id=?", (job[0]["id"],)).fetchone()
        rec = c.execute("SELECT * FROM recipients WHERE id=?", (job[1]["id"],)).fetchone()
        prev = c.execute("SELECT subject FROM messages WHERE recipient_id=? ORDER BY id DESC LIMIT 1",
                         (rec["id"],)).fetchone()
        tpl = pick_tpl(rec, job[2], s)
        subj, body = render(tpl, rec, s, prev["subject"] if prev else "")
        msgid = f"<am-{int(time.time())}-{acc['id']}-{rec['id']}@{DOMAIN}>"
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = c.execute(
            "INSERT INTO messages(account_id,recipient_id,rnd,tpl,subject,message_id,status,sent_at)"
            " VALUES(?,?,?,?,?,?,?,?)",
            (acc["id"], rec["id"], job[2], tpl, subj, msgid, "sending", now))
        mid = cur.lastrowid
        try:
            smtp_send(acc["email"], acc["pwd"], rec["email"], subj, body, msgid)
            c.execute("UPDATE messages SET status='sent' WHERE id=?", (mid,))
            c.execute("UPDATE recipients SET round_done=?, last_sent_at=?, account_id=?, status='sent' WHERE id=?",
                      (job[2], now, acc["id"], rec["id"]))
            c.execute("UPDATE accounts SET sent_total=sent_total+1, sent_today=sent_today+1, last_send_at=? WHERE id=?",
                      (now, acc["id"]))
            return True, f"{acc['email']} -> {rec['email']} [{tpl}] r{job[2]}"
        except smtplib.SMTPRecipientsRefused as e:
            c.execute("UPDATE messages SET status='bounced', error=? WHERE id=?", (str(e)[:300], mid))
            c.execute("UPDATE recipients SET status='bounced', last_sent_at=? WHERE id=?", (now, rec["id"]))
            c.execute("UPDATE accounts SET last_send_at=? WHERE id=?", (now, acc["id"]))
            return False, f"退信 {rec['email']}: {str(e)[:120]}"
        except Exception as e:
            c.execute("UPDATE messages SET status='failed', error=? WHERE id=?",
                      (f"{type(e).__name__}:{e}"[:300], mid))
            c.execute("UPDATE accounts SET last_error=? WHERE id=?", (f"{type(e).__name__}:{e}"[:200], acc["id"]))
            return False, f"失败 {acc['email']} -> {rec['email']}: {type(e).__name__}: {str(e)[:120]}"


def recover_inflight(max_age_min=10):
    """重启后收尾：把上次进程留下的 sending 行判为 unknown，并占位收件人避免重发。

    返回 (unknown 数, 受影响收件人数)。仅在启动时调用一次。
    """
    cutoff = (datetime.now() - timedelta(minutes=max_age_min)).strftime("%Y-%m-%d %H:%M:%S")
    n = m = 0
    with conn() as c:
        rows = c.execute("SELECT * FROM messages WHERE status='sending' AND sent_at <= ?", (cutoff,)).fetchall()
        for r in rows:
            c.execute("UPDATE messages SET status='unknown', error='进程中断，投递结果待人工确认' WHERE id=?", (r["id"],))
            c.execute("""UPDATE recipients SET round_done=?, last_sent_at=?, status='unconfirmed'
                         WHERE id=? AND round_done < ?""", (r["rnd"], r["sent_at"], r["recipient_id"], r["rnd"]))
            n += 1
            m += 1
    if n:
        event("recover", f"中断恢复：{n} 封状态未知，已占位不重发")
        log(f"RECOVER inflight unknown={n}")
    return n, m


def campaign_progress():
    """任务进度：今日两段（首轮 / 跟进）的目标、完成、可发与预计剩余时间"""
    s = get_settings()
    today = date.today().isoformat()
    cool = int(s.get("cooldown_min") or 15) * 60
    now = datetime.now()
    with conn() as c:
        n_acc = c.execute("SELECT COUNT(*) FROM accounts WHERE status='ok'").fetchone()[0]
        rows = c.execute("""SELECT rnd, status, COUNT(*) c FROM messages
                            WHERE sent_at LIKE ? GROUP BY rnd, status""", (today + "%",)).fetchall()
        cap = int(s.get("daily_cap") or 0)
        fcap = int(s.get("followup_cap") or 0)
        rounds = int(s.get("rounds") or 1)
        # 可跟进对象：首轮已发、非 A 类、未退信、且已过跟进间隔
        gap = int(s.get("followup_gap_hours") or 24)
        cutoff = (now - timedelta(hours=gap)).strftime("%Y-%m-%d %H:%M:%S")
        fup_pool = c.execute("""SELECT COUNT(*) FROM recipients
            WHERE round_done=1 AND status NOT IN ('bounced','unsubscribed','done','A')
              AND IFNULL(category,'')!='A' AND IFNULL(last_sent_at,'')<=?""", (cutoff,)).fetchone()[0]
        new_pool = c.execute("""SELECT COUNT(*) FROM recipients
            WHERE round_done=0 AND status NOT IN ('bounced','unsubscribed','done')""").fetchone()[0]
        accs = c.execute("SELECT last_send_at FROM accounts WHERE status='ok'").fetchall()
        unknown = c.execute("SELECT COUNT(*) FROM messages WHERE status='unknown'").fetchone()[0]

    done = {"1": 0, "2": 0}
    bad = {"1": 0, "2": 0}
    for r in rows:
        key = str(r["rnd"])
        if r["status"] != "sending":          # 进度按尝试数计，退信也算已处理
            done[key] = done.get(key, 0) + r["c"]
        if r["status"] in ("bounced", "failed"):
            bad[key] = bad.get(key, 0) + r["c"]

    # 同域日限会挡住一部分新对象：目标按"实际可发"收敛，避免进度长期停在中间
    maxdom = max(1, int(s.get("max_per_domain_day") or 1))
    with conn() as c:
        domcnt = {r["domain"]: r["c"] for r in c.execute(
            """SELECT rc.domain domain, COUNT(*) c FROM messages m
               JOIN recipients rc ON rc.id=m.recipient_id
               WHERE m.sent_at LIKE ? GROUP BY rc.domain""", (today + "%",)).fetchall()}
        sendable = 0
        for r in c.execute("""SELECT domain FROM recipients
                WHERE round_done=0 AND status NOT IN ('bounced','unsubscribed','done')"""):
            if domcnt.get(r["domain"], 0) < maxdom:
                sendable += 1
    capped = max(0, new_pool - sendable)
    d1, d2 = done.get("1", 0), done.get("2", 0)
    t1 = max(d1, min(n_acc * cap, d1 + sendable))
    t2 = min(n_acc * fcap, fup_pool) if rounds >= 2 else 0
    total_target = t1 + t2
    total_done = d1 + d2

    ready = cooling = 0
    for a in accs:
        if not a["last_send_at"]:
            ready += 1
            continue
        try:
            last = datetime.strptime(a["last_send_at"], "%Y-%m-%d %H:%M:%S")
            if (now - last).total_seconds() >= cool:
                ready += 1
            else:
                cooling += 1
        except Exception:
            ready += 1
    remain = max(0, total_target - total_done)
    per_min = (n_acc / (cool / 60.0)) if cool else float(n_acc)
    eta = int(round(remain / per_min)) if per_min > 0 else 0
    last_tick = s.get("last_tick") or ""
    stale = False
    if last_tick:
        try:
            stale = (now - datetime.strptime(last_tick, "%Y-%m-%d %H:%M:%S")).total_seconds() > 180
        except Exception:
            stale = False
    return {
        "state": s.get("campaign_state"),
        "window_open": send_window_open(s),
        "today": {"target": total_target, "done": total_done,
                  "pct": round(total_done * 100.0 / total_target, 1) if total_target else 0,
                  "remain": remain, "eta_min": eta,
                  "sent": sum(done.values()), "bad": sum(bad.values())},
        "round1": {"target": t1, "done": d1, "bad": bad.get("1", 0),
                   "pct": round(d1 * 100.0 / t1, 1) if t1 else 0},
        "round2": {"target": t2, "done": d2, "bad": bad.get("2", 0), "pool": fup_pool,
                   "pct": round(d2 * 100.0 / t2, 1) if t2 else 0},
        "queue": {"new": new_pool, "followup": fup_pool, "sendable": sendable, "capped": capped},
        "accounts": {"ready": ready, "cooling": cooling, "total": n_acc},
        "unknown": unknown,
        "last_tick": last_tick,
        "scheduler_stale": stale,
    }


def build_jobs(force=False, limit=0):
    """作业分配：先满足首轮配额（每账号每日 N 封新对象），再进入跟进段

    - 首轮(rnd=1)：收 round_done=0 的新收件人，每账号每日上限 = daily_cap
    - 跟进(rnd=2)：只追 round_done=1 且分类为 B/C/未回复者，A 类直接跳过
    - force=True 供面板预览使用：跳过状态/窗口/冷却闸门，不写任何发信记录
    """
    s = get_settings()
    if not force:
        if s.get("campaign_state") != "running":
            return []
        # 闸门一：发送时段
        if not send_window_open(s):
            return []
        # 闸门二：退信率自动刹车
        if bounce_guard(s)[0]:
            return []
    cooldown = int(s["cooldown_min"]) * 60
    cap = int(s["daily_cap"])
    fcap = int(s.get("followup_cap", "2"))
    rounds = int(s["rounds"])
    maxdom = max(1, int(s.get("max_per_domain_day") or 1))
    jobs = []
    taken = set()          # 本批次已分配收件人，避免同批重复投递
    batch_dom = {}         # 本批次各域名已用次数，用于跨域分散
    today = date.today().isoformat()
    with conn() as c:
        c.execute("UPDATE accounts SET sent_today=0, reset_date=? WHERE reset_date IS NULL OR reset_date!=?",
                  (today, today))
        # 当日各收件域名已投递数，用于同域配额
        dom_today = {r["domain"]: r["c"] for r in c.execute(
            """SELECT rc.domain domain, COUNT(*) c FROM messages m
               JOIN recipients rc ON rc.id=m.recipient_id
               WHERE m.sent_at LIKE ? GROUP BY rc.domain""", (today + "%",)).fetchall()}
        accs = c.execute("SELECT * FROM accounts WHERE status='ok' ORDER BY sent_today, last_send_at").fetchall()
        for acc in accs:
            if acc["sent_today"] >= cap + fcap:
                continue
            if acc["last_send_at"]:
                try:
                    last = datetime.strptime(acc["last_send_at"], "%Y-%m-%d %H:%M:%S")
                    if (datetime.now() - last).total_seconds() < cooldown:
                        continue
                except Exception:
                    pass
            n_init = c.execute("SELECT COUNT(*) FROM messages WHERE account_id=? AND rnd=1 AND sent_at LIKE ?",
                               (acc["id"], today + "%")).fetchone()[0]
            n_fup = c.execute("SELECT COUNT(*) FROM messages WHERE account_id=? AND rnd=2 AND sent_at LIKE ?",
                              (acc["id"], today + "%")).fetchone()[0]

            plan = []
            if n_init < cap:
                plan.append(1)
            if n_fup < fcap and rounds >= 2:
                plan.append(2)
            for rnd in plan:
                cands = [r for r in due_recipients(c, rnd, account_id=acc["id"])
                         if r["id"] not in taken
                         and dom_today.get(r["domain"], 0) + batch_dom.get(r["domain"], 0) < maxdom]
                if not cands:
                    continue
                # 优先：本账号未发过 > 本批未用过该域 > 全局被发次数少 > 有姓名
                def score(r):
                    used = c.execute("SELECT COUNT(*) FROM recipients WHERE domain=? AND account_id=?",
                                     (r["domain"], acc["id"])).fetchone()[0]
                    dcount = c.execute("SELECT COUNT(*) FROM recipients WHERE domain=? AND round_done>0",
                                       (r["domain"],)).fetchone()[0]
                    return (used, batch_dom.get(r["domain"], 0), dcount, 0 if r["name"] else 1)
                cands = sorted(cands, key=score)
                pick = cands[0]
                taken.add(pick["id"])
                batch_dom[pick["domain"]] = batch_dom.get(pick["domain"], 0) + 1
                jobs.append((acc, pick, rnd))
                break
    return jobs


class Scheduler(threading.Thread):
    def __init__(self):
        super().__init__(daemon=True)
        self.stop_flag = False

    def run(self):
        log("scheduler started")
        while not self.stop_flag:
            try:
                jobs = build_jobs()
                gap = max(2.0, float(get_settings().get("send_gap_sec") or 20))
                for j in jobs:
                    ok, msg = send_one(j)
                    log(("SENT " if ok else "FAIL ") + msg)
                    time.sleep(gap * random.uniform(0.6, 1.4))
                if jobs:
                    event("tick", f"本轮发出 {len(jobs)} 封")
                set_settings({"last_tick": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
            except Exception:
                log("scheduler error:", traceback.format_exc()[:800])
            time.sleep(12)


# ---------------------------------------------------------------- 收信 + 分类
def decode_mbox(name: str) -> str:
    """IMAP modified UTF-7 -> str"""
    try:
        if "&" in name:
            name = re.sub(r"&([A-Za-z0-9+,]*)-", lambda m: "+" + m.group(1).replace(",", "/") + "-", name)
            return name.encode().decode("utf-7")
    except Exception:
        pass
    return name


BOUNCE_SUBJ = re.compile(r"(undelivered|delivery status|mail delivery|returned mail|退信|投递失败|发送失败|failure notice)", re.I)
BOUNCE_FROM = re.compile(r"(mailer-daemon|postmaster|mail-daemon)", re.I)
SYSTEM_SENDER = re.compile(r"(@mailsupport\.aliyun\.com|@mail\.aliyun\.com|@mxhichina\.com|@service\.aliyun\.com)", re.I)
AUTO_SUBJ = re.compile(r"(auto(matic)?[ -]?reply|out of (the )?office|autoreply|自动回复|休假|不在办公室)", re.I)

QUOTE_WORDS = re.compile(r"(price|pricing|rate|rates|cost|quote|quotation|fee|budget|USD|EUR|GBP|\$|€|£|报价|价格|费用|美元|单价)", re.I)
METRIC_WORDS = re.compile(r"\b(DA|DR|Ahrefs|Moz|traffic|visitors|sessions)\b|流量", re.I)
ACCEPT_WORDS = re.compile(r"(available|availability|accept|we can|we do|sure|yes|order|publish|insert|sponsor|可以|接受|能做|可做|合作)", re.I)
REJECT_WORDS = re.compile(r"(not interested|no interest|unsubscribe|remove (me|us)|opt[ -]?out|stop (emailing|sending)|do not contact|don'?t contact|no thanks|not relevant|不感兴趣|不需要|不要联系|退订|拒绝|暂不)", re.I)
MONEY = re.compile(r"(?:USD|EUR|GBP|\$|€|£)\s?\d|(?:\d{2,6})\s?(?:USD|EUR|GBP|usd|eur|gbp)|(?:\d{2,6})\s?(?:美元|元|欧元)")


def classify(subject: str, body: str, from_addr: str):
    """返回 (category, auto, bounce, quote_hit, snippet)"""
    text = f"{subject}\n{body}"
    if BOUNCE_FROM.search(from_addr or "") or BOUNCE_SUBJ.search(subject or ""):
        return "X", 0, 1, 0, body.strip()[:200]
    auto = 1 if AUTO_SUBJ.search(subject or "") else 0
    if REJECT_WORDS.search(text):
        cat = "C"
    elif MONEY.search(text) or (QUOTE_WORDS.search(text) and (
            re.search(r"\d{2,}", text) or METRIC_WORDS.search(text))):
        cat = "A"
    elif QUOTE_WORDS.search(text) and ACCEPT_WORDS.search(text):
        cat = "A"
    else:
        cat = "B"
    quote_hit = 1 if (MONEY.search(text) or (QUOTE_WORDS.search(text) and re.search(r"\d{2,}", text))) else 0
    snippet = re.sub(r"\s+", " ", text).strip()[:220]
    return cat, auto, 0, quote_hit, snippet


URL_RE = re.compile(r"https?://[^\s<>\"'\)\]]+", re.I)
SHORTENERS = {"bit.ly", "t.co", "t.cn", "ow.ly", "goo.gl", "tinyurl.com", "is.gd", "buff.ly",
              "rebrand.ly", "cutt.ly", "shorturl.at", "rb.gy", "s.id", "v.gd"}
BAD_TLD = {"zip", "mov", "top", "xyz", "click", "link", "work", "kim", "gq", "tk", "ml", "ga", "cf",
           "buzz", "cyou", "rest", "sbs", "cam", "monster"}


def extract_links(text):
    seen, out = set(), []
    for u in URL_RE.findall(text or ""):
        u = u.rstrip(".,;:!?)]}\"'")
        if u.lower() in seen:
            continue
        seen.add(u.lower())
        out.append(u)
    return out[:20]


def check_link(url):
    flags, code, final = [], None, ""
    try:
        p = urlparse(url)
        host = (p.hostname or "").lower()
        if p.scheme not in ("http", "https"):
            flags.append("non-http")
        if not host:
            return None, "", ["no-host"]
        if re.fullmatch(r"\d{1,3}(\.\d{1,3}){3}", host):
            flags.append("ip-literal")
        if host in SHORTENERS:
            flags.append("shortener")
        if host.startswith("xn--") or ".xn--" in host:
            flags.append("punycode")
        tld = host.rsplit(".", 1)[-1] if "." in host else ""
        if tld in BAD_TLD:
            flags.append("suspicious-tld:" + tld)
        if p.port and p.port not in (80, 443):
            flags.append("odd-port:" + str(p.port))
        try:
            socket.getaddrinfo(host, None)
        except Exception:
            flags.append("dns-fail")
        req = urllib.request.Request(url, method="HEAD",
                                     headers={"User-Agent": "Mozilla/5.0 (AutomailLinkCheck)"})
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        try:
            with urllib.request.urlopen(req, timeout=8, context=ctx) as r:
                code, final = r.status, r.geturl()
        except urllib.error.HTTPError as e:
            code, final = e.code, url
        except Exception as e:
            flags.append("fetch-fail:" + type(e).__name__)
        if code and code >= 400:
            flags.append(f"http-{code}")
        if final and urlparse(final).hostname and urlparse(final).hostname != host:
            flags.append("redirect:" + (urlparse(final).hostname or ""))
    except Exception as e:
        flags.append("parse-fail:" + type(e).__name__)
    return code, final, flags


def get_text(msg):
    plain, html = "", ""
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        try:
            if part.get_content_type() == "text/plain":
                plain += part.get_payload(decode=True).decode(part.get_content_charset() or "utf8", "ignore")
            elif part.get_content_type() == "text/html":
                html += part.get_payload(decode=True).decode(part.get_content_charset() or "utf8", "ignore")
        except Exception:
            continue
    if html:
        txt = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", html, flags=re.S | re.I)
        txt = re.sub(r"<[^>]+>", " ", txt)
        txt = re.sub(r"&nbsp;?", " ", txt)
        txt = re.sub(r"\s+", " ", txt)
        if len(txt) > len(plain):
            plain = txt
    return plain


def record_inbound(acct, msg, raw_body):
    """把一封收进来的邮件入库 + 分类 + 链接检查"""
    subject = str(msg.get("Subject", ""))
    from_addr = parseaddr(str(msg.get("From", "")))[1].lower()
    mid_hdr = str(msg.get("In-Reply-To", "")) or str(msg.get("References", ""))
    cat, auto, bounce, quote_hit, snippet = classify(subject, raw_body, from_addr)
    with conn() as c:
        sent = None
        for part in (mid_hdr or "").split():
            p = part.strip("<> ")
            sent = c.execute("SELECT * FROM messages WHERE message_id LIKE ?", (f"%{p}%",)).fetchone()
            if sent:
                break
        rec = c.execute("SELECT * FROM recipients WHERE id=?", (sent["recipient_id"],)).fetchone() if sent else None
        if rec is None:
            rec = c.execute("SELECT * FROM recipients WHERE email=?", (from_addr,)).fetchone()
        links = extract_links(raw_body)
        cur = c.execute(
            "INSERT INTO replies(account_id,message_id,from_addr,subject,received_at,body,category,auto,is_bounce,links_json,quote_hit,snippet)"
            " VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
            (acct["id"], sent["message_id"] if sent else "", from_addr, subject,
             datetime.now().strftime("%Y-%m-%d %H:%M:%S"), raw_body[:20000], cat, auto, bounce,
             json.dumps(links, ensure_ascii=False), quote_hit, snippet))
        rid = cur.lastrowid
        if bounce:
            # 从退信正文里识别是哪个收件人
            for cand in c.execute("SELECT id,email FROM recipients WHERE status!='bounced'").fetchall():
                if cand["email"] in raw_body:
                    c.execute("UPDATE recipients SET status='bounced' WHERE id=?", (cand["id"],))
                    c.execute("UPDATE messages SET status='bounced' WHERE recipient_id=?", (cand["id"],))
                    break
        elif rec is not None:
            if cat == "C":
                c.execute("UPDATE recipients SET category=?, reply_at=?, status='replied' WHERE id=?",
                          (cat, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), rec["id"]))
            elif cat == "A":
                c.execute("UPDATE recipients SET category=?, reply_at=?, status='A' WHERE id=?",
                          (cat, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), rec["id"]))
            else:
                c.execute("UPDATE recipients SET category=?, reply_at=?, status='replied' WHERE id=?",
                          (cat, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), rec["id"]))
    # 链接检查（同步，限 20 条）；退信通知里的厂商链接无分析价值，跳过
    check_on = get_settings().get("link_check") == "1"
    for u in ([] if bounce else links):
        if check_on:
            code, final, flags = check_link(u)
        else:
            code, final, flags = None, "", ["unchecked"]
        with conn() as c:
            c.execute("INSERT INTO link_checks(reply_id,url,status_code,final_url,flags,checked_at)"
                      " VALUES(?,?,?,?,?,?)",
                      (rid, u, code, final, json.dumps(flags, ensure_ascii=False),
                       datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    event("inbound", f"{acct['email']} <- {from_addr} [{cat}] {subject[:60]}")
    return cat, from_addr


def poll_account(acc, first_run_days=2):
    s = get_settings()
    ctx = ssl.create_default_context()
    got = 0
    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ctx) as m:
        m.login(acc["email"], acc["pwd"])
        typ, boxes = m.list()
        names = []
        for b in boxes or []:
            try:
                nm = b.decode("utf8", "ignore").split(' "/" ')[-1].strip('"')
            except Exception:
                continue
            names.append(nm)
        targets = ["INBOX"]
        if s.get("scan_spam") == "1":
            targets += [n for n in names if re.search(r"(junk|spam|垃圾)", decode_mbox(n), re.I)]
        for folder in targets:
            try:
                typ, _ = m.select(f'"{folder}"', readonly=True)
                if typ != "OK":
                    continue
                with conn() as c:
                    fr = c.execute("SELECT last_uid FROM folders WHERE account_id=? AND folder=?",
                                   (acc["id"], folder)).fetchone()
                last = fr["last_uid"] if fr else 0
                if last:
                    typ, data = m.uid("SEARCH", None, f"UID {last+1}:*")
                else:
                    since = (date.today() - timedelta(days=first_run_days)).strftime("%d-%b-%Y")
                    typ, data = m.uid("SEARCH", None, f'SINCE {since}')
                uids = [int(x) for x in (data[0].split() if data and data[0] else [])]
                uids = [u for u in uids if u > last]
                for uid in uids[-40:]:
                    typ, d = m.uid("FETCH", str(uid), "(RFC822)")
                    if typ != "OK":
                        continue
                    raw = next((x[1] for x in d if isinstance(x, tuple)), None)
                    if not raw:
                        continue
                    msg = message_from_bytes(raw, policy=policy.default)
                    frm = parseaddr(str(msg.get("From", "")))[1].lower()
                    if frm.endswith("@" + DOMAIN) or SYSTEM_SENDER.search(frm):
                        last = max(last, uid)
                        continue
                    try:
                        body = get_text(msg)
                        record_inbound(acc, msg, body)
                        got += 1
                    except Exception:
                        log("inbound error", traceback.format_exc()[:400])
                    last = max(last, uid)
                with conn() as c:
                    c.execute("INSERT INTO folders(account_id,folder,last_uid) VALUES(?,?,?)"
                              " ON CONFLICT(account_id,folder) DO UPDATE SET last_uid=excluded.last_uid",
                              (acc["id"], folder, last))
            except Exception:
                log("folder error", folder, traceback.format_exc()[:300])
    return got


class Poller(threading.Thread):
    def __init__(self):
        super().__init__(daemon=True)

    def run(self):
        log("poller started")
        time.sleep(10)
        while True:
            try:
                s = get_settings()
                with conn() as c:
                    accs = c.execute("SELECT * FROM accounts WHERE status='ok'").fetchall()
                total = 0
                for acc in accs:
                    try:
                        total += poll_account(acc)
                    except Exception:
                        log("poll fail", acc["email"], traceback.format_exc()[:300])
                if total:
                    log("poller new inbound:", total)
                time.sleep(max(60, int(s.get("poll_sec", 180))))
            except Exception:
                log("poller error", traceback.format_exc()[:500])
                time.sleep(60)


# ---------------------------------------------------------------- 导出 A 类
def export_a():
    """导出 A 类（有意向报价）客户清单：以回信为准，附报价金额、判定依据与回信内链接"""
    os.makedirs(EXPORT_DIR, exist_ok=True)
    price_re = re.compile(r"(?:[€$£]\s?\d[\d.,]*|\d[\d.,]*\s?(?:euros?|eur|usd|dollars?|gbp|元))", re.I)
    rows = q("""SELECT rp.id, rp.from_addr, rp.subject, rp.received_at, rp.body, rp.snippet,
                       rp.links_json, rp.ai_json, rp.quote_hit,
                       rc.domain, rc.name, rc.company
                FROM replies rp LEFT JOIN recipients rc ON LOWER(rc.email)=LOWER(rp.from_addr)
                WHERE IFNULL(rp.is_bounce,0)=0 AND rp.category='A'
                ORDER BY rp.received_at DESC""")
    for r in rows:
        body = (r.pop("body") or "")
        price_hits = {m.group(0).strip() for m in price_re.finditer(body)}

        def _pv(x):
            try:
                return float(re.sub(r"[^0-9.]", "", x) or 0)
            except Exception:
                return 0.0
        r["prices"] = sorted(price_hits, key=_pv)[:8]
        try:
            ai = json.loads(r.pop("ai_json") or "{}")
        except Exception:
            ai = {}
        r["quote"] = ai.get("quote") or ""
        r["reason"] = ai.get("reason") or ""
        try:
            r["links"] = json.loads(r.pop("links_json") or "[]")
        except Exception:
            r["links"] = []
        r["domain"] = r["domain"] or (r["from_addr"].split("@")[-1] if "@" in r["from_addr"] else "")
        # 正文里提到的其它可投站点；排除社交/图床/邮箱等非售卖域名
        skip = {"facebook.com", "twitter.com", "x.com", "instagram.com", "linkedin.com",
                "youtube.com", "youtu.be", "pinterest.com", "tumblr.com", "blogspot.com",
                "wordpress.com", "medium.com", "gmail.com", "google.com", "whatsapp.com",
                "t.me", "telegram.me", "gravatar.com", "wixsite.com", "weebly.com"}
        seen = set()
        for d in re.findall(
                r"\b((?:[a-z0-9-]+\.)+(?:com|net|org|fr|de|es|it|pl|be|se|br|info|co|io|nl|pt|gr|hu|ro))\b",
                body, re.I):
            d = d.lower()
            if d == r["domain"] or d in skip or any(d.endswith("." + x) for x in skip):
                continue
            seen.add(d)
        r["sites"] = sorted(seen)[:12]

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    head = f"{datetime.now():%Y-%m-%d %H:%M}"
    md = [f"# A 类客户（已报价） {head}", "",
          f"共 {len(rows)} 封报价回信。判定口径：回信给出明确单价或报价区间，可直接进入议价/下单。", ""]
    if not rows:
        md.append("_暂无 A 类回信_")
    for r in rows:
        md += [f"## {r['domain']}",
               f"- 邮箱：{r['from_addr']}",
               f"- 联系人：{r['name'] or '-'}　公司：{r['company'] or '-'}",
               f"- 回信时间：{r['received_at'] or '-'}",
               f"- 主题：{r['subject'] or '-'}",
               f"- 报价：{'、'.join(r['prices']) if r['prices'] else '未识别到金额，见正文'}",
               f"- 判定依据：{r['reason'] or r['quote'] or '正文命中报价特征'}"]
        if r["sites"]:
            md.append(f"- 可投站点：{', '.join(r['sites'])}")
        if r["links"]:
            md.append("- 回信内链接：")
            md += [f"  - {u}" for u in r["links"]]
        md.append("")

    esc = lambda x: (str(x or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    html = ["""<meta charset="utf-8"><title>A 类客户清单</title>
<style>
body{font-family:-apple-system,"Microsoft YaHei",sans-serif;margin:32px;color:#1b1c1f;background:#fff;line-height:1.65}
h1{font-size:21px;margin:0 0 6px}.sub{color:#6a6b73;font-size:12.5px;margin-bottom:20px}
table{border-collapse:collapse;width:100%;font-size:12.5px}
th,td{border:1px solid #e2e4e8;padding:9px 10px;vertical-align:top;text-align:left}
th{background:#f7f8fa;font-weight:600;white-space:nowrap}
.price{color:#a8372c;font-weight:600}code{background:#f5f6f8;padding:1px 4px;border-radius:3px;font-size:11.5px}
.empty{color:#8a8b93;padding:18px 0}
</style>""",
            f"<h1>A 类客户（已报价）· {head}</h1>",
            f'<div class="sub">共 <b>{len(rows)}</b> 封报价回信。判定口径：回信给出明确单价或报价区间，可直接进入议价/下单。</div>']
    if rows:
        html.append("<table><tr><th>域名</th><th>邮箱 / 联系人</th><th>回信时间</th><th>报价</th>"
                    "<th>判定依据</th><th>可投站点</th><th>回信内链接</th></tr>")
        for r in rows:
            html.append("<tr>"
                        f"<td><b>{esc(r['domain'])}</b></td>"
                        f"<td>{esc(r['from_addr'])}<br><span style='color:#8a8b93'>{esc(r['name'] or '-')} · {esc(r['company'] or '-')}</span></td>"
                        f"<td>{esc(r['received_at'])}<br><span style='color:#8a8b93'>{esc((r['subject'] or '')[:46])}</span></td>"
                        f"<td class='price'>{esc('、'.join(r['prices']) or '见正文')}</td>"
                        f"<td>{esc(r['reason'] or r['quote'] or '正文命中报价特征')}</td>"
                        f"<td>{esc(', '.join(r['sites']) or '-')}</td>"
                        f"<td>{'<br>'.join('<code>' + esc(u) + '</code>' for u in r['links']) or '-'}</td>"
                        "</tr>")
        html.append("</table>")
    else:
        html.append('<div class="empty">暂无 A 类回信</div>')

    p_md = os.path.join(EXPORT_DIR, f"A类客户_{stamp}.md")
    p_html = os.path.join(EXPORT_DIR, f"A类客户_{stamp}.html")
    open(p_md, "w", encoding="utf-8").write("\n".join(md))
    open(p_html, "w", encoding="utf-8").write("\n".join(html))
    event("export_a", f"A 类清单导出 {len(rows)} 封报价回信")
    return p_md, p_html, len(rows)


# ---------------------------------------------------------------- 表格导入
ALIAS = {
    "email": ["邮箱", "邮件地址", "email", "e-mail", "mail", "账号", "邮箱账号", "收件人邮箱"],
    "password": ["密码", "password", "pwd", "邮箱密码"],
    "name": ["姓名", "联系人", "名字", "name", "contact", "first name"],
    "company": ["公司名称", "公司", "company", "organization"],
    "domain": ["域名", "domain", "网站", "website"],
}


def _norm(h):
    return str(h or "").strip().lower().replace("（", "(").replace("）", ")")


def parse_table(filename, raw):
    """csv / xlsx -> list[dict]（按别名映射 email/password/name/company/domain）"""
    name = (filename or "").lower()
    rows = []
    if name.endswith((".xlsx", ".xlsm")):
        import io
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        ws = wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        header = [_norm(x) for x in (next(it, None) or [])]
        for r in it:
            rows.append(dict(zip(header, ["" if v is None else str(v).strip() for v in r])))
    else:
        text = None
        for enc in ("utf-8-sig", "gbk", "utf-8"):
            try:
                text = raw.decode(enc)
                break
            except Exception:
                continue
        if text is None:
            text = raw.decode("utf-8", "ignore")
        rd = csv.reader(text.splitlines())
        header = [_norm(x) for x in (next(rd, None) or [])]
        for r in rd:
            rows.append(dict(zip(header, [str(x).strip() for x in r])))

    def field(row, key):
        """按 精确 > 前缀 > 包含 三级匹配；并排除跨字段误命中（如密码列的说明文字含"邮箱账号"）"""
        bad = ("密码", "password", "pwd") if key == "email" else ("email", "邮件地址")
        def clean(k):
            return not any(b in k for b in bad)
        cands = [(k, v) for k, v in row.items() if v and clean(k)]
        for a in ALIAS[key]:
            na = _norm(a)
            for k, v in cands:
                if k == na:
                    return v
            for k, v in cands:
                if k.startswith(na):
                    return v
        for a in ALIAS[key]:
            na = _norm(a)
            for k, v in cands:
                if na in k:
                    return v
        return ""

    out = []
    for row in rows:
        email = field(row, "email").lower().strip().strip("<>").split()[-1] if field(row, "email") else ""
        if "@" not in email or "." not in email.split("@")[-1]:
            continue
        dom = field(row, "domain").lower().strip() or email.split("@")[1]
        dom = re.sub(r"^https?://|^www\.|/.*$", "", dom)
        out.append({"email": email, "password": field(row, "password"),
                    "name": field(row, "name"), "company": field(row, "company"), "domain": dom})
    return out


def import_accounts(filename, raw):
    rows = parse_table(filename, raw)
    ok = dup = 0
    with conn() as c:
        for r in rows:
            pwd = r["password"] or PASSWORD
            cur = c.execute("INSERT OR IGNORE INTO accounts(email,pwd,reset_date) VALUES(?,?,?)",
                            (r["email"], pwd, date.today().isoformat()))
            ok += cur.rowcount
            dup += 0 if cur.rowcount else 1
        total = c.execute("SELECT COUNT(*) FROM accounts").fetchone()[0]
    event("import_accounts", f"{filename}: 新增 {ok}，重复 {dup}")
    return {"added": ok, "dup": dup, "total": total}


def import_recipients(filename, raw):
    rows = parse_table(filename, raw)
    ok = dup = 0
    with conn() as c:
        for r in rows:
            cur = c.execute("INSERT OR IGNORE INTO recipients(email,name,domain,company) VALUES(?,?,?,?)",
                            (r["email"], r["name"], r["domain"], r["company"]))
            ok += cur.rowcount
            dup += 0 if cur.rowcount else 1
        total = c.execute("SELECT COUNT(*) FROM recipients").fetchone()[0]
    event("import_recipients", f"{filename}: 新增 {ok}，重复 {dup}")
    return {"added": ok, "dup": dup, "total": total}


# ---------------------------------------------------------------- 域名池
DOMAIN_RE = re.compile(r"\b((?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+(?:[a-z]{2,24}))\b", re.I)
NON_DOMAIN_TLD = {"jpg", "jpeg", "png", "gif", "webp", "svg", "css", "js", "json", "pdf", "zip",
                  "doc", "docx", "xls", "xlsx", "mp4", "mp3", "rar", "html", "php", "aspx"}


def _cell_texts(filename, raw):
    """把上传文件摊平成待扫描的文本片段"""
    name = (filename or "").lower()
    if name.endswith((".xlsx", ".xlsm")):
        import io
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        out = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if v is not None:
                        out.append(str(v))
        return out
    text = None
    for enc in ("utf-8-sig", "gbk", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except Exception:
            continue
    if text is None:
        text = raw.decode("utf-8", "ignore")
    out = []
    for line in text.splitlines():
        out.extend(x.strip() for x in re.split(r"[,;\t]", line) if x.strip())
    return out


def extract_domains(filename, raw):
    """从 csv / xlsx / txt 抽取域名；单元格里是网址或邮箱也能取到主机名"""
    seen, out = set(), []
    for frag in _cell_texts(filename, raw):
        f = frag.strip().strip('"\'<>')
        if not f:
            continue
        cands = []
        if "://" in f:
            try:
                cands.append(urlparse(f).hostname or "")
            except Exception:
                pass
        for m in DOMAIN_RE.findall(f):
            cands.append(m)
        for d in cands:
            d = (d or "").lower().strip(".").strip()
            d = re.sub(r"^www\.", "", d)
            if not d or "@" in d or d == DOMAIN:
                continue
            tld = d.rsplit(".", 1)[-1]
            if tld in NON_DOMAIN_TLD or not re.fullmatch(r"[a-z0-9.-]{4,253}", d):
                continue
            if d not in seen:
                seen.add(d)
                out.append(d)
    return out


def import_domains(filename, raw, expand=None, parts=None):
    """上传域名名单。expand=True 时按常用前缀展开为候选邮箱入池"""
    s = get_settings()
    doms = extract_domains(filename, raw)
    if expand is None:
        expand = str(s.get("expand_on_upload", "1")) == "1"
    plist = parts if parts is not None else s.get("domain_local_parts", "info,contact,sales")
    plist = [re.sub(r"[^a-z0-9._-]", "", p.strip().lower()) for p in str(plist).split(",")]
    plist = [p for p in plist if p][:20]
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_dom = addr = 0
    with conn() as c:
        for d in doms:
            cur = c.execute("INSERT OR IGNORE INTO domain_pool(domain,source,added_at) VALUES(?,?,?)",
                            (d, filename, now))
            new_dom += cur.rowcount
            row = c.execute("SELECT * FROM domain_pool WHERE domain=?", (d,)).fetchone()
            if not expand or (row and row["expanded"]):
                continue
            added = 0
            for p in plist:
                rc = c.execute("INSERT OR IGNORE INTO recipients(email,name,domain,company) VALUES(?,?,?,?)",
                               (f"{p}@{d}", "", d, ""))
                added += rc.rowcount
            addr += added
            c.execute("UPDATE domain_pool SET expanded=1, addr_added=? WHERE domain=?", (added, d))
        total_dom = c.execute("SELECT COUNT(*) FROM domain_pool").fetchone()[0]
        total_rec = c.execute("SELECT COUNT(*) FROM recipients").fetchone()[0]
    event("import_domains", f"{filename}: 域名 {len(doms)}（新增 {new_dom}），展开新增邮箱 {addr}")
    return {"domains": len(doms), "new_domains": new_dom, "added": addr,
            "total_domains": total_dom, "total": total_rec}


def expand_domains(domains=None, parts=None):
    """把已登记但未展开的域名补展开为候选邮箱"""
    s = get_settings()
    plist = parts if parts is not None else s.get("domain_local_parts", "info,contact,sales")
    plist = [re.sub(r"[^a-z0-9._-]", "", p.strip().lower()) for p in str(plist).split(",")]
    plist = [p for p in plist if p][:20]
    ids = []
    with conn() as c:
        if domains:
            for d in domains:
                row = c.execute("SELECT id FROM domain_pool WHERE domain=?", (str(d).lower().strip(),)).fetchone()
                if row:
                    ids.append(row["id"])
        else:
            ids = [r["id"] for r in c.execute("SELECT id FROM domain_pool WHERE expanded=0").fetchall()]
        added = 0
        for i in ids:
            row = c.execute("SELECT * FROM domain_pool WHERE id=?", (i,)).fetchone()
            if not row:
                continue
            n = 0
            for p in plist:
                rc = c.execute("INSERT OR IGNORE INTO recipients(email,name,domain,company) VALUES(?,?,?,?)",
                               (f"{p}@{row['domain']}", "", row["domain"], ""))
                n += rc.rowcount
            added += n
            c.execute("UPDATE domain_pool SET expanded=1, addr_added=? WHERE id=?", (n, i))
        total = c.execute("SELECT COUNT(*) FROM recipients").fetchone()[0]
    event("expand_domains", f"展开 {len(ids)} 个域名，新增邮箱 {added}")
    return {"domains": len(ids), "added": added, "total": total}


def export_daily(dstr=None):
    """导出某日收发明细 xlsx：发信 / 回信 / 账号汇总 / 退信"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    dstr = dstr or date.today().isoformat()
    os.makedirs(EXPORT_DIR, exist_ok=True)
    wb = openpyxl.Workbook()
    head_font = Font(bold=True)
    fill = PatternFill("solid", fgColor="F2F2F2")

    def sheet(ws, headers, rows, widths=None):
        ws.append(headers)
        for c in ws[1]:
            c.font = head_font
            c.fill = fill
            c.alignment = Alignment(vertical="center")
        for r in rows:
            ws.append([r.get(h, "") if isinstance(r, dict) else r for h in headers] if rows and isinstance(rows[0], dict) else r)
        for i, h in enumerate(headers, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = (
                widths[i - 1] if widths else max(12, min(46, len(str(h)) * 2 + 4)))

    sent = q("""SELECT m.sent_at 时间, a.email 发件账号, r.email 收件人, r.domain 域名, m.rnd 轮次,
                m.tpl 模板, m.subject 主题, m.status 状态, m.error 错误
                FROM messages m LEFT JOIN recipients r ON r.id=m.recipient_id
                LEFT JOIN accounts a ON a.id=m.account_id
                WHERE m.sent_at LIKE ? ORDER BY m.id""", (dstr + "%",))
    sheet(wb.active, list(sent[0].keys()) if sent else ["时间", "发件账号", "收件人", "域名", "轮次", "模板", "主题", "状态", "错误"],
          sent)
    wb.active.title = f"发信_{dstr}"

    rep = q("""SELECT rp.received_at 时间, rp.from_addr 来源, rp.category 分类, rp.auto 自动回复,
               rp.is_bounce 退信, a.email 收件账号, rp.subject 主题, rp.snippet 摘要, rp.ai_done AI已处理
               FROM replies rp LEFT JOIN accounts a ON a.id=rp.account_id
               WHERE rp.received_at LIKE ? ORDER BY rp.id""", (dstr + "%",))
    sheet(wb.create_sheet(f"回信_{dstr}"),
          list(rep[0].keys()) if rep else ["时间", "来源", "分类", "自动回复", "退信", "收件账号", "主题", "摘要", "AI已处理"], rep)

    accs = q("""SELECT a.email 账号, a.sent_today 今日已发, a.sent_total 累计已发,
                (SELECT COUNT(*) FROM messages m WHERE m.account_id=a.id AND m.status='bounced') 退信,
                (SELECT COUNT(*) FROM messages m WHERE m.account_id=a.id AND m.sent_at LIKE ?) 当日发出,
                a.last_send_at 最近发信, a.last_error 最近错误
                FROM accounts a ORDER BY a.email""", (dstr + "%",))
    sheet(wb.create_sheet("账号汇总"), list(accs[0].keys()), accs)

    bad = q("""SELECT m.sent_at 时间, a.email 发件账号, r.email 收件人, m.status 状态, m.error 错误
               FROM messages m LEFT JOIN accounts a ON a.id=m.account_id
               LEFT JOIN recipients r ON r.id=m.recipient_id
               WHERE m.status!='sent' ORDER BY m.id DESC""")
    sheet(wb.create_sheet("退信与失败"),
          list(bad[0].keys()) if bad else ["时间", "发件账号", "收件人", "状态", "错误"], bad)

    path = os.path.join(EXPORT_DIR, f"Automail日志_{dstr}.xlsx")
    wb.save(path)
    return path


def export_daily_csv(dstr=None):
    """当日收发日志导出为单份 CSV（多段拼接，Excel 可直接打开）"""
    dstr = dstr or date.today().isoformat()
    os.makedirs(EXPORT_DIR, exist_ok=True)
    sent = q("""SELECT m.sent_at, a.email acc, r.email rec, r.domain, m.rnd, m.tpl, m.subject, m.status, m.error
                FROM messages m LEFT JOIN recipients r ON r.id=m.recipient_id
                LEFT JOIN accounts a ON a.id=m.account_id
                WHERE m.sent_at LIKE ? ORDER BY m.id""", (dstr + "%",))
    rep = q("""SELECT rp.received_at, rp.from_addr, rp.category, rp.auto, rp.is_bounce, a.email acc,
               rp.subject, rp.snippet
               FROM replies rp LEFT JOIN accounts a ON a.id=rp.account_id
               WHERE rp.received_at LIKE ? ORDER BY rp.id""", (dstr + "%",))
    bad = q("""SELECT m.sent_at, a.email acc, r.email rec, m.status, m.error
               FROM messages m LEFT JOIN accounts a ON a.id=m.account_id
               LEFT JOIN recipients r ON r.id=m.recipient_id
               WHERE m.status!='sent' AND m.sent_at LIKE ? ORDER BY m.id""", (dstr + "%",))
    accs = q("""SELECT a.email, a.sent_total, a.last_send_at, a.last_error,
                (SELECT COUNT(*) FROM messages m WHERE m.account_id=a.id AND m.sent_at LIKE ?) today,
                (SELECT COUNT(*) FROM messages m WHERE m.account_id=a.id AND m.status='bounced') bounced
                FROM accounts a ORDER BY a.email""", (dstr + "%",))
    path = os.path.join(EXPORT_DIR, f"Automail日志_{dstr}.csv")
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)

        def block(title, headers, rows, keys):
            w.writerow([f"## {title}（{len(rows)} 条）"])
            w.writerow(headers)
            for r in rows:
                w.writerow([r.get(k, "") for k in keys])
            w.writerow([])

        block("发信明细", ["时间", "发件账号", "收件人", "域名", "轮次", "模板", "主题", "状态", "错误"],
              sent, ["sent_at", "acc", "rec", "domain", "rnd", "tpl", "subject", "status", "error"])
        block("回信明细", ["时间", "来源", "分类", "自动回复", "退信", "收件账号", "主题", "摘要"],
              rep, ["received_at", "from_addr", "category", "auto", "is_bounce", "acc", "subject", "snippet"])
        block("退信与失败", ["时间", "发件账号", "收件人", "状态", "错误"],
              bad, ["sent_at", "acc", "rec", "status", "error"])
        block("账号汇总", ["账号", "累计已发", "当日发出", "退信", "最近发信", "最近错误"],
              accs, ["email", "sent_total", "today", "bounced", "last_send_at", "last_error"])
        st = today_stats()
        w.writerow(["## 当日汇总"])
        w.writerow(["发出成功", "失败或退信", "退信率%", "导出时间"])
        w.writerow([st["sent"], st["bad"], st["rate"], datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    return path


def list_exports():
    os.makedirs(EXPORT_DIR, exist_ok=True)
    out = []
    for n in sorted(os.listdir(EXPORT_DIR), reverse=True):
        p = os.path.join(EXPORT_DIR, n)
        if not os.path.isfile(p):
            continue
        out.append({"name": n, "size": os.path.getsize(p),
                    "mtime": datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y-%m-%d %H:%M:%S"),
                    "ext": os.path.splitext(n)[1].lstrip(".")})
    return out[:100]


# ---------------------------------------------------------------- 看板聚合
def _pct(a, b):
    return round(a * 100.0 / b, 1) if b else 0.0


def dashboard():
    """总览看板一次性聚合：KPI / 进度 / 构成 / 实时动态 / 账号矩阵"""
    s = get_settings()
    today = date.today().isoformat()
    now = datetime.now()
    prog = campaign_progress()
    cool = int(s.get("cooldown_min") or 15) * 60

    with conn() as c:
        t_rows = c.execute("""SELECT rnd, status, COUNT(*) c FROM messages
                              WHERE sent_at LIKE ? GROUP BY rnd, status""", (today + "%",)).fetchall()
        all_rows = c.execute("SELECT status, COUNT(*) c FROM messages GROUP BY status").fetchall()
        accs = c.execute("""SELECT id, email, status, sent_total, sent_today, last_send_at, last_error
                            FROM accounts ORDER BY email""").fetchall()
        acc_today = {r["aid"]: r for r in c.execute(
            """SELECT account_id aid, COUNT(*) c, SUM(status='sent') ok, SUM(status!='sent') bad
               FROM messages WHERE sent_at LIKE ? GROUP BY account_id""", (today + "%",)).fetchall()}
        acc_all = {r["aid"]: r for r in c.execute(
            """SELECT account_id aid, COUNT(*) c, SUM(status='sent') ok, SUM(status='bounced') bad
               FROM messages GROUP BY account_id""").fetchall()}
        feed = c.execute("""SELECT m.id, m.sent_at, m.status, m.tpl, m.rnd,
                                   a.email acc, rc.email rec, rc.domain rdom
                            FROM messages m LEFT JOIN accounts a ON a.id=m.account_id
                            LEFT JOIN recipients rc ON rc.id=m.recipient_id
                            ORDER BY m.id DESC LIMIT 18""").fetchall()
        by_tpl = c.execute("""SELECT tpl, SUM(status='sent') c,
                                     COUNT(DISTINCT CASE WHEN status='sent' THEN account_id END) a
                              FROM messages WHERE sent_at LIKE ? GROUP BY tpl ORDER BY c DESC""",
                           (today + "%",)).fetchall()
        by_dom = c.execute("""SELECT rc.domain d, SUM(m.status='sent') c, SUM(m.status!='sent') bad,
                                     COUNT(DISTINCT m.account_id) a
                              FROM messages m JOIN recipients rc ON rc.id=m.recipient_id
                              WHERE m.sent_at LIKE ? GROUP BY rc.domain ORDER BY c DESC""",
                           (today + "%",)).fetchall()
        dom_all = c.execute("""SELECT domain, COUNT(*) total, SUM(round_done>0) used,
                                      SUM(status='bounced') bad
                               FROM recipients GROUP BY domain ORDER BY total DESC""").fetchall()
        rep = c.execute("""SELECT COUNT(*) total, SUM(is_bounce=1) bounce,
                                  SUM(is_bounce=0 AND IFNULL(category,'')='A') a,
                                  SUM(is_bounce=0 AND IFNULL(category,'')='B') b,
                                  SUM(is_bounce=0 AND IFNULL(category,'')='C') cc,
                                  SUM(is_bounce=0 AND IFNULL(ai_done,0)=0) pend
                           FROM replies""").fetchone()
        days = c.execute("""SELECT COUNT(DISTINCT substr(sent_at,1,10)) d FROM messages
                            WHERE status='sent'""").fetchone()[0]
        hourly = [0] * 24
        for r in c.execute("""SELECT substr(sent_at,12,2) h, COUNT(*) c FROM messages
                              WHERE sent_at LIKE ? AND status='sent' GROUP BY h""", (today + "%",)).fetchall():
            try:
                hourly[int(r["h"])] = r["c"]
            except Exception:
                pass

    today_sent = sum(r["c"] for r in t_rows if r["status"] == "sent")
    today_bad = sum(r["c"] for r in t_rows if r["status"] in ("bounced", "failed"))
    all_sent = sum(r["c"] for r in all_rows if r["status"] == "sent")
    all_bad = sum(r["c"] for r in all_rows if r["status"] in ("bounced", "failed"))

    # 账号矩阵
    groups, matrix = {}, []
    for a in accs:
        local, _, dom = a["email"].partition("@")
        at = acc_today.get(a["id"])
        al = acc_all.get(a["id"])
        t_ok = (at["ok"] or 0) if at else 0
        t_bad = (at["bad"] or 0) if at else 0
        a_ok = (al["ok"] or 0) if al else 0
        a_bad = (al["bad"] or 0) if al else 0
        cooling = 0
        if a["last_send_at"]:
            try:
                cooling = max(0, int(cool - (now - datetime.strptime(
                    a["last_send_at"], "%Y-%m-%d %H:%M:%S")).total_seconds()))
            except Exception:
                cooling = 0
        if a["status"] != "ok":
            state = "stopped"
        elif cooling > 0:
            state = "cooling"
        elif t_ok + t_bad == 0:
            state = "idle"
        else:
            state = "active"
        rate = round(a_ok * 100.0 / (a_ok + a_bad), 1) if (a_ok + a_bad) else 0.0
        item = {"email": a["email"], "local": local, "domain": dom, "state": state,
                "today": t_ok, "today_bad": t_bad, "total": a_ok, "bad": a_bad,
                "rate": rate, "last": a["last_send_at"] or "",
                "cool_left": int(cooling // 60) if cooling else 0,
                "next": (datetime.strptime(a["last_send_at"], "%Y-%m-%d %H:%M:%S")
                         + timedelta(seconds=cool)).strftime("%H:%M")
                        if a["last_send_at"] else "",
                "error": a["last_error"] or ""}
        matrix.append(item)
        groups.setdefault(dom, []).append(item)

    grp = [{"domain": d, "n": len(v),
            "today": sum(x["today"] for x in v),
            "bad": sum(x["bad"] for x in v),
            "stopped": sum(1 for x in v if x["state"] == "stopped"),
            "accounts": v} for d, v in groups.items()]
    grp.sort(key=lambda g: -g["n"])

    total_sent = prog["today"]["done"]
    target = prog["today"]["target"]
    pending = prog["queue"]["sendable"] + prog["queue"]["followup"]

    return {
        "kpi": {
            "today_sent": today_sent, "today_target": target,
            "today_bad": today_bad,
            "pending": pending, "capped": prog["queue"]["capped"],
            "acc_ready": prog["accounts"]["ready"], "acc_total": prog["accounts"]["total"],
            "acc_cooling": prog["accounts"]["cooling"],
            "total_sent": all_sent, "avg_daily": int(round(all_sent / days)) if days else 0,
            "delivery": round(all_sent * 100.0 / (all_sent + all_bad), 1) if (all_sent + all_bad) else 0.0,
            "bounced": all_bad,
        },
        "progress": {
            "done": total_sent, "target": target, "pct": prog["today"]["pct"],
            "remain": prog["today"]["remain"], "eta": prog["today"].get("eta_min") or 0,
            "accounts_ok": prog["accounts"]["ready"] + prog["accounts"]["cooling"],
            "accounts_total": prog["accounts"]["total"],
            "pct_acc": round((prog["accounts"]["ready"] + prog["accounts"]["cooling"]) * 100.0
                             / (prog["accounts"]["total"] or 1), 1),
            "r1": prog["round1"], "r2": prog["round2"], "queue": prog["queue"],
            "unknown": prog["unknown"],
        },
        "feed": [{"id": r["id"], "t": (r["sent_at"] or "")[11:19], "d": (r["sent_at"] or "")[5:10],
                  "acc": (r["acc"] or "").split("@")[0], "rec": r["rec"] or "",
                  "rdom": r["rdom"] or "", "status": r["status"], "tpl": r["tpl"], "rnd": r["rnd"]}
                 for r in feed],
        "composition": {
            "tpl": [{"k": r["tpl"] or "—", "n": r["c"], "a": r["a"]} for r in by_tpl],
            "domain": [{"k": r["d"], "n": r["c"], "bad": r["bad"] or 0, "a": r["a"]} for r in by_dom],
        },
        "hourly": hourly,
        "domains": [{"domain": r["domain"], "total": r["total"], "used": r["used"] or 0,
                     "bad": r["bad"] or 0} for r in dom_all[:40]],
        "replies": {"total": rep["total"] or 0, "bounce": rep["bounce"] or 0,
                    "A": rep["a"] or 0, "B": rep["b"] or 0, "C": rep["cc"] or 0,
                    "pending": rep["pend"] or 0},
        "groups": grp,
        "state": prog["state"], "window_open": prog["window_open"],
        "window": f"{s.get('send_window_start')}–{s.get('send_window_end')}",
        "last_tick": prog["last_tick"], "stale": prog.get("scheduler_stale", False),
        "now": now.strftime("%H:%M:%S"),
    }


def account_detail(email):
    with conn() as c:
        a = c.execute("SELECT * FROM accounts WHERE email=?", (email,)).fetchone()
        if not a:
            return {"error": "账号不存在"}
        today = date.today().isoformat()
        t = c.execute("""SELECT COUNT(*) c, SUM(status='sent') ok, SUM(status!='sent') bad
                         FROM messages WHERE account_id=? AND sent_at LIKE ?""",
                      (a["id"], today + "%")).fetchone()
        al = c.execute("""SELECT COUNT(*) c, SUM(status='sent') ok, SUM(status='bounced') bad
                          FROM messages WHERE account_id=?""", (a["id"],)).fetchone()
        msgs = c.execute("""SELECT m.id, m.sent_at, m.status, m.tpl, m.rnd, m.subject, m.error,
                                   rc.email rec, rc.domain rdom, rc.category
                            FROM messages m LEFT JOIN recipients rc ON rc.id=m.recipient_id
                            WHERE m.account_id=? ORDER BY m.id DESC LIMIT 12""", (a["id"],)).fetchall()
        rep = c.execute("""SELECT COUNT(*) c FROM replies WHERE account_id=? AND is_bounce=0""",
                        (a["id"],)).fetchone()[0]
        days = c.execute("""SELECT COUNT(DISTINCT substr(sent_at,1,10)) d FROM messages
                            WHERE account_id=? AND status='sent'""", (a["id"],)).fetchone()[0]
    ok = al["ok"] or 0
    bad = al["bad"] or 0
    s = get_settings()
    cool = int(s.get("cooldown_min") or 15)
    nxt = ""
    if a["last_send_at"]:
        try:
            nxt = (datetime.strptime(a["last_send_at"], "%Y-%m-%d %H:%M:%S")
                   + timedelta(minutes=cool)).strftime("%H:%M")
        except Exception:
            nxt = ""
    return {"email": a["email"], "status": a["status"], "today": t["ok"] or 0,
            "today_bad": t["bad"] or 0, "total": ok, "bad": bad,
            "rate": round(ok * 100.0 / (ok + bad), 1) if (ok + bad) else 0.0,
            "avg": int(round(ok / days)) if days else 0,
            "last": a["last_send_at"] or "", "next": nxt, "replies": rep,
            "error": a["last_error"] or "",
            "msgs": [dict(m) for m in msgs]}


# ---------------------------------------------------------------- API
def q(sql, args=()):
    with conn() as c:
        return [dict(r) for r in c.execute(sql, args).fetchall()]


class Handler(BaseHTTPRequestHandler):
    server_version = "AutomailDesk/1.0"

    def _send(self, code, body=b"", ctype="application/json; charset=utf-8"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8"))

    def body_json(self):
        n = int(self.headers.get("Content-Length") or 0)
        if not n:
            return {}
        try:
            return json.loads(self.rfile.read(n).decode("utf-8"))
        except Exception:
            return {}

    def do_GET(self):
        p = urlparse(self.path).path
        try:
            if p in ("/", "/index.html"):
                with open(UI, "rb") as f:
                    return self._send(200, f.read(), "text/html; charset=utf-8")
            if p == "/api/overview":
                today = date.today().isoformat()
                d = {
                    "today_sent": q("SELECT COUNT(*) c FROM messages WHERE status='sent' AND sent_at LIKE ?", (today + "%",))[0]["c"],
                    "total_sent": q("SELECT COUNT(*) c FROM messages WHERE status='sent'")[0]["c"],
                    "bounced": q("SELECT COUNT(*) c FROM messages WHERE status='bounced'")[0]["c"],
                    "replies": q("SELECT COUNT(*) c FROM replies WHERE is_bounce=0")[0]["c"],
                    "A": q("SELECT COUNT(*) c FROM recipients WHERE category='A'")[0]["c"],
                    "B": q("SELECT COUNT(*) c FROM recipients WHERE category='B'")[0]["c"],
                    "C": q("SELECT COUNT(*) c FROM recipients WHERE category='C'")[0]["c"],
                    "accounts": q("SELECT COUNT(*) c FROM accounts")[0]["c"],
                    "recipients": q("SELECT COUNT(*) c FROM recipients")[0]["c"],
                    "round1": q("SELECT COUNT(*) c FROM recipients WHERE round_done>=1")[0]["c"],
                    "round2": q("SELECT COUNT(*) c FROM recipients WHERE round_done>=2")[0]["c"],
                    "round3": q("SELECT COUNT(*) c FROM recipients WHERE round_done>=3")[0]["c"],
                    "links": q("SELECT COUNT(*) c FROM link_checks")[0]["c"],
                    "link_warn": q("SELECT COUNT(*) c FROM link_checks WHERE flags NOT LIKE '[]'")[0]["c"],
                    "state": get_settings().get("campaign_state"),
                    "window_open": send_window_open(),
                    "window": f"{get_settings().get('send_window_start')}–{get_settings().get('send_window_end')}",
                    "bounce_rate": today_stats()["rate"],
                    "domains": q("SELECT COUNT(*) c FROM domain_pool")[0]["c"],
                    "ai_pending": q("SELECT COUNT(*) c FROM replies WHERE IFNULL(ai_done,0)=0 AND is_bounce=0")[0]["c"],
                }
                return self.json(d)
            if p == "/api/trend":
                out = []
                for i in range(6, -1, -1):
                    d = (date.today() - timedelta(days=i)).isoformat()
                    out.append({
                        "date": d,
                        "sent": q("SELECT COUNT(*) c FROM messages WHERE sent_at LIKE ? AND status='sent'", (d + "%",))[0]["c"],
                        "bad": q("SELECT COUNT(*) c FROM messages WHERE sent_at LIKE ? AND status!='sent'", (d + "%",))[0]["c"],
                        "replies": q("SELECT COUNT(*) c FROM replies WHERE received_at LIKE ? AND is_bounce=0", (d + "%",))[0]["c"],
                    })
                return self.json(out)
            if p == "/api/accounts":
                return self.json(q("""SELECT a.*, 
                    (SELECT COUNT(*) FROM messages m WHERE m.account_id=a.id AND m.status='sent') ok,
                    (SELECT COUNT(*) FROM messages m WHERE m.account_id=a.id AND m.status='bounced') bad,
                    (SELECT MAX(sent_at) FROM messages m WHERE m.account_id=a.id) last
                    FROM accounts a ORDER BY a.email"""))
            if p == "/api/unknown":
                return self.json(q("""SELECT m.id, m.sent_at, m.rnd, m.status,
                        a.email acc_email, r.email rec_email, r.domain
                    FROM messages m LEFT JOIN recipients r ON r.id=m.recipient_id
                    LEFT JOIN accounts a ON a.id=m.account_id
                    WHERE m.status='unknown' ORDER BY m.id DESC LIMIT 200"""))
            if p == "/api/messages":
                lim = 500
                return self.json(q("""SELECT m.*, r.email rec_email, r.domain, a.email acc_email
                    FROM messages m LEFT JOIN recipients r ON r.id=m.recipient_id
                    LEFT JOIN accounts a ON a.id=m.account_id
                    ORDER BY m.id DESC LIMIT ?""", (lim,)))
            if p == "/api/replies":
                return self.json(q("""SELECT rp.id,rp.from_addr,rp.subject,rp.received_at,rp.category,rp.auto,
                    rp.is_bounce,rp.snippet,rp.quote_hit,rp.ai_done,rp.ai_json,a.email acc_email
                    FROM replies rp LEFT JOIN accounts a ON a.id=rp.account_id
                    ORDER BY rp.id DESC LIMIT 500"""))
            if p == "/api/ai/pending":
                rows = q("""SELECT rp.id,rp.from_addr,rp.subject,rp.received_at,rp.category,rp.auto,
                    rp.is_bounce,rp.body,rp.snippet,rp.links_json,a.email acc_email
                    FROM replies rp LEFT JOIN accounts a ON a.id=rp.account_id
                    WHERE IFNULL(rp.ai_done,0)=0 AND rp.is_bounce=0 ORDER BY rp.id DESC LIMIT 50""")
                return self.json({"count": len(rows), "rows": rows})
            if p == "/api/export/daily":
                qs = urlparse(self.path).query
                d, fmt = "", "xlsx"
                for kv in qs.split("&"):
                    if kv.startswith("date="):
                        d = unquote(kv[5:])
                    elif kv.startswith("fmt="):
                        fmt = unquote(kv[4:]).lower()
                d = d or date.today().isoformat()
                path = export_daily_csv(d) if fmt == "csv" else export_daily(d)
                return self.json({"path": path, "file": os.path.basename(path),
                                  "url": "/dl/" + os.path.basename(path)})
            if p == "/api/exports":
                return self.json(list_exports())
            if p.startswith("/dl/"):
                name = unquote(p[4:])
                fp = os.path.join(EXPORT_DIR, os.path.basename(name))
                if not os.path.isfile(fp):
                    return self._send(404, b'{"error":"file not found"}')
                ctype = ("text/csv; charset=utf-8" if fp.lower().endswith(".csv")
                         else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                         if fp.lower().endswith(".xlsx") else "application/octet-stream")
                with open(fp, "rb") as f:
                    return self._send(200, f.read(), ctype)
            if p == "/api/preflight":
                return self.json(preflight())
            if p == "/api/progress":
                return self.json(campaign_progress())
            if p == "/api/dashboard":
                return self.json(dashboard())
            if p == "/api/account_detail":
                qs = urlparse(self.path).query
                email = ""
                for kv in qs.split("&"):
                    if kv.startswith("email="):
                        email = unquote(kv[6:])
                return self.json(account_detail(email))
            if p == "/api/domains":
                # 域名池 = 已登记域名 ∪ 收件人名单里的域名，避免上传名单后此页空白
                pool = {r["domain"]: r for r in q("SELECT * FROM domain_pool")}
                out = []
                for r in q("""SELECT domain, COUNT(*) addr, SUM(round_done>0) sent,
                                     SUM(status='bounced') bad
                              FROM recipients GROUP BY domain"""):
                    meta = pool.pop(r["domain"], None)
                    out.append({"domain": r["domain"],
                                "source": (meta["source"] if meta else "收件人名单"),
                                "added_at": (meta["added_at"] if meta else ""),
                                "expanded": 1,
                                "addr": r["addr"], "sent": r["sent"] or 0,
                                "bad": r["bad"] or 0})
                for d, meta in pool.items():
                    out.append({"domain": d, "source": meta["source"], "added_at": meta["added_at"],
                                "expanded": meta["expanded"], "addr": 0, "sent": 0, "bad": 0})
                out.sort(key=lambda x: (-x["addr"], x["domain"]))
                return self.json(out[:500])
            if p == "/api/ai/tasks":
                qs = urlparse(self.path).query
                lim = 20
                for kv in qs.split("&"):
                    if kv.startswith("limit="):
                        try:
                            lim = max(1, min(100, int(kv[6:])))
                        except Exception:
                            pass
                rows = q("""SELECT rp.id, rp.from_addr, rp.subject, rp.received_at, rp.category,
                            rp.snippet, rp.body, rp.links_json, rp.quote_hit, a.email acc_email
                            FROM replies rp LEFT JOIN accounts a ON a.id=rp.account_id
                            WHERE IFNULL(rp.ai_done,0)=0 AND rp.is_bounce=0
                            ORDER BY rp.id DESC LIMIT ?""", (lim,))
                for r in rows:
                    body = (r.pop("body") or "")[:4000]
                    r["body"] = body
                    try:
                        r["links"] = json.loads(r.pop("links_json") or "[]")
                    except Exception:
                        r["links"] = []
                return self.json({"count": len(rows), "rows": rows,
                                  "howto": "POST /api/ai/result {reply_id,category,note,quote?,reason?}，category 取 A/B/C"})
            if p == "/api/preview":
                s = get_settings()
                jobs = build_jobs(force=True, limit=12)
                rows = [{"account": a["email"], "recipient": r["email"], "rnd": rnd,
                         "tpl": pick_tpl(r, rnd, s)} for a, r, rnd in jobs]
                note = ""
                if not send_window_open(s):
                    note = f"当前 {datetime.now():%H:%M} 在窗口 {s.get('send_window_start')}–{s.get('send_window_end')} 之外，实际调度会挂起"
                elif s.get("campaign_state") != "running":
                    note = "任务未启动，以下是按参数推演的分配结果"
                elif not rows:
                    note = "队列为空或今日配额已用尽"
                return self.json({"n": len(rows), "rows": rows, "note": note})
            if p == "/api/events":
                return self.json(q("SELECT * FROM events ORDER BY id DESC LIMIT 60"))
            if p.startswith("/api/reply/"):
                rid = int(p.rsplit("/", 1)[-1])
                d = q("SELECT * FROM replies WHERE id=?", (rid,))
                if not d:
                    return self.json({"error": "not found"}, 404)
                rel = dict(d[0])
                rel["links"] = q("SELECT * FROM link_checks WHERE reply_id=?", (rid,))
                return self.json(rel)
            if p == "/api/links":
                return self.json(q("""SELECT lc.*, rp.from_addr FROM link_checks lc
                    LEFT JOIN replies rp ON rp.id=lc.reply_id
                    WHERE IFNULL(rp.is_bounce,0)=0
                    ORDER BY lc.id DESC LIMIT 400"""))
            if p == "/api/settings":
                return self.json(get_settings())
            if p == "/api/recipients":
                return self.json(q("""SELECT * FROM recipients ORDER BY
                    CASE category WHEN 'A' THEN 0 WHEN 'B' THEN 1 WHEN 'C' THEN 2 ELSE 3 END, id LIMIT 400"""))
            if p == "/api/export_a":
                md, html, n = export_a()
                return self.json({"md": md, "html": html, "count": n})
            if p == "/api/poll":
                with conn() as c:
                    accs = c.execute("SELECT * FROM accounts WHERE status='ok'").fetchall()
                n = 0
                for a in accs:
                    try:
                        n += poll_account(a)
                    except Exception:
                        pass
                return self.json({"new": n})
            return self._send(404, b'{"error":"not found"}')
        except Exception:
            log("api error", traceback.format_exc()[:600])
            return self.json({"error": traceback.format_exc()[-400:]}, 500)

    def do_POST(self):
        p = urlparse(self.path).path
        try:
            if p == "/api/settings":
                patch = self.body_json()
                set_settings(patch)
                event("settings", json.dumps({k: str(v)[:60] for k, v in patch.items()}, ensure_ascii=False))
                return self.json({"ok": True, "settings": get_settings()})
            if p == "/api/unknown":
                b = self.body_json() or {}
                try:
                    mid = int(b.get("id") or 0)
                except Exception:
                    return self.json({"error": "bad id"}, 400)
                act = b.get("action")
                if act not in ("sent", "retry"):
                    return self.json({"error": "bad action"}, 400)
                with conn() as c:
                    m = c.execute("SELECT * FROM messages WHERE id=? AND status='unknown'", (mid,)).fetchone()
                    if not m:
                        return self.json({"error": "not found"}, 404)
                    if act == "sent":
                        c.execute("UPDATE messages SET status='sent', error='人工确认已投递' WHERE id=?", (mid,))
                        c.execute("UPDATE recipients SET status='sent' WHERE id=?", (m["recipient_id"],))
                    else:
                        c.execute("UPDATE messages SET status='void', error='人工标记重发' WHERE id=?", (mid,))
                        c.execute("""UPDATE recipients SET round_done=?, status='pending', last_sent_at=NULL
                                     WHERE id=?""", (max(0, int(m["rnd"] or 1) - 1), m["recipient_id"]))
                event("unknown", f"人工处理未确认 #{mid} -> {act}")
                return self.json({"ok": True, "progress": campaign_progress()})
            if p == "/api/shutdown":
                # 面板右上角停止按钮：回执先写回，随后结束整个进程（调度/收信线程随进程退出）
                event("shutdown", "面板停止按钮：关闭调度、收信与面板服务")
                self.json({"ok": True})

                def _die():
                    time.sleep(0.5)
                    log("SHUTDOWN by panel button")
                    os._exit(0)

                threading.Thread(target=_die, daemon=True).start()
                return
            if p == "/api/resume":
                if READONLY:
                    return self.json({"error": "readonly sandbox"}, 403)
                b = self.body_json() or {}
                rec, _ = recover_inflight(2 if not b.get("force") else 0)
                set_settings({"campaign_state": "running",
                              "last_tick": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                event("resume", f"手动续发，收尾未知状态 {rec} 封")
                jobs = build_jobs()
                return self.json({"ok": True, "recovered": rec, "jobs": len(jobs),
                                  "progress": campaign_progress()})
            if p == "/api/campaign":
                st = (self.body_json() or {}).get("state")
                if st not in ("running", "paused", "idle"):
                    return self.json({"error": "bad state"}, 400)
                set_settings({"campaign_state": st})
                event("campaign", st)
                return self.json({"ok": True, "state": st})
            if p.startswith("/api/reply/") and p.endswith("/category"):
                rid = int(p.split("/")[3])
                cat = (self.body_json() or {}).get("category", "")
                if cat not in ("A", "B", "C", "X", ""):
                    return self.json({"error": "bad category"}, 400)
                with conn() as c:
                    r = c.execute("SELECT from_addr FROM replies WHERE id=?", (rid,)).fetchone()
                    c.execute("UPDATE replies SET category=? WHERE id=?", (cat, rid))
                    if r:
                        c.execute("UPDATE recipients SET category=?, status=? WHERE email=?",
                                  (cat, "A" if cat == "A" else "replied", r["from_addr"]))
                return self.json({"ok": True})
            if p == "/api/send_now":
                if READONLY:
                    return self.json({"error": "readonly sandbox"}, 403)
                jobs = build_jobs()
                res = []
                for j in jobs:
                    ok, msg = send_one(j)
                    res.append({"ok": ok, "msg": msg})
                return self.json({"n": len(res), "result": res})
            if p == "/api/render":
                b = self.body_json() or {}
                s = get_settings()
                key = b.get("tpl") or "S1"
                tpls = dict(s.get("templates") or {})
                tpls[key] = {"subject": b.get("subject") or "", "body": b.get("body") or ""}
                s["templates"] = tpls
                rec = {"name": "Ahmet", "domain": "example.com", "company": "Example Media"}
                subj, body = render(key, rec, s, "")
                residue = sorted(set(re.findall(r"\{[a-z_]+\}", subj + body)))
                return self.json({"subject": subj, "body": body, "residue": residue})
            if p == "/api/import/domains":
                n = int(self.headers.get("Content-Length") or 0)
                raw = self.rfile.read(n) if n else b""
                fname = unquote(self.headers.get("X-Filename") or "domains.csv")
                if not raw:
                    return self.json({"error": "空文件"}, 400)
                expand, parts = None, None
                for kv in urlparse(self.path).query.split("&"):
                    if kv.startswith("expand="):
                        expand = kv[7:] == "1"
                    elif kv.startswith("parts="):
                        parts = unquote(kv[6:])
                return self.json({"ok": True, "file": fname, **import_domains(fname, raw, expand, parts)})
            if p == "/api/templates/reset":
                set_settings({"templates": DEFAULT_TEMPLATES})
                event("templates", "模板已恢复为内置默认")
                return self.json({"ok": True, "templates": get_settings().get("templates")})
            if p == "/api/domains/expand":
                b = self.body_json() or {}
                return self.json({"ok": True, **expand_domains(b.get("domains") or None, b.get("parts") or None)})
            if p == "/api/import/accounts" or p == "/api/import/recipients":
                n = int(self.headers.get("Content-Length") or 0)
                raw = self.rfile.read(n) if n else b""
                fname = unquote(self.headers.get("X-Filename") or "upload.csv")
                if not raw:
                    return self.json({"error": "空文件"}, 400)
                fn = import_accounts if p.endswith("accounts") else import_recipients
                return self.json({"ok": True, "file": fname, **fn(fname, raw)})
            if p == "/api/ai/result":
                b = self.body_json()
                items = b.get("items") if isinstance(b.get("items"), list) else [b]
                done = 0
                for it in items:
                    try:
                        rid = int(it.get("reply_id") or 0)
                    except Exception:
                        continue
                    if not rid:
                        continue
                    cat = it.get("category") or ""
                    with conn() as c:
                        r = c.execute("SELECT from_addr FROM replies WHERE id=?", (rid,)).fetchone()
                        c.execute("UPDATE replies SET ai_done=1, ai_json=? WHERE id=?",
                                  (json.dumps(it, ensure_ascii=False), rid))
                        if cat in ("A", "B", "C") and r:
                            c.execute("UPDATE replies SET category=? WHERE id=?", (cat, rid))
                            c.execute("UPDATE recipients SET category=?, status=? WHERE email=?",
                                      (cat, "A" if cat == "A" else "replied", r["from_addr"]))
                    done += 1
                return self.json({"ok": True, "updated": done})
            if p == "/api/replies/ignore":
                b = self.body_json()
                rid = int(b.get("reply_id") or 0)
                with conn() as c:
                    c.execute("UPDATE replies SET ai_done=1, category='I' WHERE id=?", (rid,))
                return self.json({"ok": True})
            return self._send(404, b'{"error":"not found"}')
        except Exception:
            log("post error", traceback.format_exc()[:600])
            return self.json({"error": traceback.format_exc()[-400:]}, 500)

    def log_message(self, fmt, *args):
        pass


class PanelServer(ThreadingHTTPServer):
    """禁止端口复用：避免旧进程未退时新进程也监听，请求被两套代码轮流处理"""
    allow_reuse_address = False
    daemon_threads = True


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "seed":
        return seed()
    init_db()
    if not READONLY:
        recover_inflight()
        Scheduler().start()
        Poller().start()
    try:
        srv = PanelServer(("127.0.0.1", PORT), Handler)
    except OSError as e:
        log(f"端口 {PORT} 已被占用，未启动新实例：{e}")
        print(f"端口 {PORT} 已被占用，请先关闭已有面板窗口再启动。")
        return
    log(f"panel on http://127.0.0.1:{PORT}")
    print("=" * 56)
    print("  Automail Desk  http://127.0.0.1:%d" % PORT)
    print("  Base    : %s" % BASE)
    print("  Domain  : %s" % DOMAIN)
    print("  Accounts: %d configured" % ACCOUNT_COUNT)
    if not PASSWORD or DOMAIN == "example.com":
        print("  [i] 未配置账号池：请复制 config.example.json 为 config.json 后填写")
    print("=" * 56, flush=True)
    srv.serve_forever()


if __name__ == "__main__":
    main()
