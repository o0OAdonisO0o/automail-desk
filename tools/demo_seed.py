# -*- coding: utf-8 -*-
"""生成演示数据（全部为虚构域名与地址），用于截图与界面预览，不涉及真实邮箱。

用法：
    python tools/demo_seed.py            # 默认写入 ./_demo
    python tools/demo_seed.py <base>     # 指定目录
然后：
    AUTOMAIL_BASE=<base> AUTOMAIL_PORT=8767 AUTOMAIL_READONLY=1 python server.py
"""
import os
import random
import sys
from datetime import date, datetime, timedelta

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASE = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "_demo"))
os.environ["AUTOMAIL_BASE"] = BASE
os.environ["AUTOMAIL_READONLY"] = "1"
sys.path.insert(0, HERE)

import server  # noqa: E402

random.seed(7)

ACCOUNTS = ["demo%03d@demo-mail.test" % i for i in range(1, 13)]
DOMAINS = [
    "northfieldmedia.com", "brightlanepress.com", "harborviewdaily.com", "weblinejournal.com",
    "southernpost.net", "mapleleafreview.com", "urbanvoice.co", "thelakeview.org",
    "trailheadmag.com", "greenhalldaily.com", "sunsetwire.com", "orchidpress.net",
]
LOCALS = ["editor", "hello", "sales", "info", "contact", "press"]
NAMES = ["Emma Carter", "Liam Novak", "Sofia Rossi", "Noah Fischer", "Mia Laurent",
         "Lucas Meyer", "Ava Silva", "Ethan Brooks", "Chloe Dubois", "Mason Reid",
         "Isla Turner", "Leo Martin"]
TPLS = ["S1", "S2", "F1", "F2"]


def main():
    server.init_db()
    with server.conn() as c:
        for t in ("accounts", "recipients", "messages", "replies", "link_checks",
                  "folders", "domain_pool", "settings", "events"):
            c.execute("DROP TABLE IF EXISTS %s" % t)
    server.init_db()
    with server.conn() as c:
        now = datetime.now()
        # 收件人池：12 个域名 × 3-4 个地址
        rid = 1
        recs = []
        for d in DOMAINS:
            for local in random.sample(LOCALS, random.choice([3, 3, 4])):
                email = "%s@%s" % (local, d)
                c.execute("""INSERT OR IGNORE INTO recipients
                             (email,name,domain,company,source_id,status,round_done,last_sent_at,account_id,auto)
                             VALUES(?,?,?,?,?,?,?,?,?,1)""",
                          (email, random.choice(NAMES), d, d.split(".")[0].title(),
                           "demo-%03d" % rid, "new", 0, None, None))
                recs.append((rid, email, d))
                rid += 1

        # 账号池
        aid = 1
        for e in ACCOUNTS:
            c.execute("INSERT OR IGNORE INTO accounts(email,pwd,reset_date,status) VALUES(?,?,?, 'ok')",
                      (e, "demo", date.today().isoformat()))
            aid += 1

        # 三天发信记录：前两天满量，今天进行到约 60%
        msg_id = 1
        today = date.today()
        for day_off in (2, 1, 0):
            day = today - timedelta(days=day_off)
            quota = 36 if day_off else 22
            picked = random.sample(recs, min(quota, len(recs)))
            for i, (rid_, email, dom) in enumerate(picked):
                acc = ACCOUNTS[i % len(ACCOUNTS)]
                acc_id = ACCOUNTS.index(acc) + 1
                if day_off == 0:
                    start = datetime(day.year, day.month, day.day, 9, 2)
                    end = max(start + timedelta(minutes=30), datetime.now() - timedelta(minutes=6))
                    ts = start + timedelta(seconds=random.randint(0, int((end - start).total_seconds())))
                else:
                    hour = 9 + (i * 11 // max(1, quota)) % 10
                    ts = datetime(day.year, day.month, day.day, hour, random.randint(0, 59),
                                  random.randint(0, 59))
                status = "bounced" if random.random() < 0.07 else "sent"
                tpl = random.choice(TPLS[:2] if day_off == 0 else TPLS)
                c.execute("""INSERT INTO messages(account_id,recipient_id,rnd,tpl,subject,message_id,status,sent_at,error)
                             VALUES(?,?,?,?,?,?,?,?,?)""",
                          (acc_id, rid_, 1 if tpl.startswith("S") else 2, tpl,
                           "Publishing rates and availability - %s" % dom,
                           "<demo-%d@demo-mail.test>" % msg_id, status,
                           ts.strftime("%Y-%m-%d %H:%M:%S"),
                           "550 mailbox unavailable" if status == "bounced" else None))
                msg_id += 1
                c.execute("UPDATE recipients SET round_done=1, last_sent_at=?, account_id=?, status=? WHERE id=?",
                          (ts.strftime("%Y-%m-%d %H:%M:%S"), acc_id,
                           "bounced" if status == "bounced" else "sent", rid_))

        # 账号当日计数与最后发送时间（矩阵据此显示活跃 / 冷却 / 停发）
        now_s = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for idx, e in enumerate(ACCOUNTS, start=1):
            row = c.execute("""SELECT COUNT(*) n, SUM(status='sent') ok,
                                      SUM(substr(sent_at,1,10)=?) tday, MAX(sent_at) last
                               FROM messages WHERE account_id=?""",
                            (today.isoformat(), idx)).fetchone()
            last = row["last"]
            if not last or last > now_s:
                last = (datetime.now() - timedelta(minutes=random.randint(3, 150))).strftime("%Y-%m-%d %H:%M:%S")
            c.execute("UPDATE accounts SET sent_total=?, sent_today=?, last_send_at=?, status=? WHERE id=?",
                      (row["n"] or 0, row["tday"] or 0, last,
                       "paused" if idx in (9,) else "ok", idx))

        # 回信
        replies = [
            ("editor@northfieldmedia.com", "Re: Publishing partnership with northfieldmedia.com",
             "B", "Thanks for reaching out. Please send your rate card and a sample of recent placements.", 1),
            ("sales@brightlanepress.com", "Re: Publishing rates and availability",
             "A", "We can offer guest posts at EUR 60 and link insertions at EUR 45. Casino content is EUR 180. "
                  "Payment via bank transfer, NET15. Live examples attached below.", 1),
            ("hello@harborviewdaily.com", "Re: Publishing partnership",
             "A", "Our rates: general EUR 95, iGaming EUR 220, CBD EUR 140. Minimum order 3 placements per quarter. "
                  "DA 58, DR 61, 120k monthly visits.", 1),
            ("contact@thelakeview.org", "Re: Publishing rates",
             "C", "We are not accepting sponsored content at the moment. Please remove us from your list.", 1),
            ("info@weblinejournal.com", "Re: Publishing partnership with weblinejournal.com",
             "B", "Not the right contact. Forwarding to our partnerships team.", 0),
            ("press@orchidpress.net", "Re: Publishing rates and availability - orchidpress.net",
             "B", "Could you share your budget range first?", 1),
        ]
        for i, (frm, subj, cat, body, done) in enumerate(replies, start=1):
            ts = (datetime.now() - timedelta(hours=20 - i * 3)).strftime("%Y-%m-%d %H:%M:%S")
            c.execute("""INSERT INTO replies(account_id,message_id,from_addr,subject,received_at,body,
                                            category,auto,is_bounce,links_json,quote_hit,snippet,ai_done)
                         VALUES(?,?,?,?,?,?,?,1,0,'[]',?,?,?)""",
                      (i, "<in-%d@demo-mail.test>" % i, frm, subj, ts, body, cat,
                       1 if "EUR" in body else 0, body[:120], done))
            c.execute("UPDATE recipients SET category=?, status=?, reply_at=? WHERE email=?",
                      (cat, "A" if cat == "A" else "replied", ts, frm))

        # 链接检查
        c.execute("""INSERT INTO link_checks(reply_id,url,status_code,final_url,flags,checked_at)
                     VALUES(1,'https://northfieldmedia.com/rates',200,'https://northfieldmedia.com/rates','',?),
                           (2,'https://brightlanepress.com/advertise',200,'https://brightlanepress.com/advertise','',?),
                           (3,'https://bit.ly/demo-rate',302,'https://harborviewdaily.com/rates','shortlink',?)""",
                  tuple([datetime.now().strftime("%Y-%m-%d %H:%M:%S")] * 3))

        # 设置
        c.execute("UPDATE settings SET v='running' WHERE k='campaign_state'")
        c.execute("UPDATE settings SET v='+1 555 0100' WHERE k='whatsapp'")
        c.execute("UPDATE settings SET v='Northwind Media' WHERE k='company_name'")
        for k, detail in (("scheduler", "调度器启动，窗口 09:00-20:00"),
                          ("send", "S1 → editor@northfieldmedia.com 已投递"),
                          ("reply", "新回信：sales@brightlanepress.com 判定 A 类")):
            c.execute("INSERT INTO events(ts,kind,detail) VALUES(?,?,?)",
                      (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), k, detail))
    print("demo db:", server.DB_PATH)


if __name__ == "__main__":
    main()
