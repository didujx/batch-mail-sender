# -*- coding: utf-8 -*-
"""
按行个性化发邮件（一人一封，正文取 B 列）—— 一次性批量发送

用法:
    pip install openpyxl            # 只有 .xlsx 才需要；.csv 不需要装任何东西
    python send_emails.py                                  # 使用同目录 config.ini + 收件人名单.xlsx
    python send_emails.py --dry-run                        # 只预演不发送
    python send_emails.py --limit 3                        # 只发前 3 条，先自己验收
    python send_emails.py --excel 名单.xlsx --sheet Sheet1
    python send_emails.py --send                           # 使用已保存的授权令牌(免密)发送

安全设计:
    1. 默认 dry-run，必须在命令里显式加 --no-dry-run 才真正发送。
    2. 每条邮件都经过校验(邮箱格式/正文为空/重复)，不合规的行直接跳过并记入日志。
    3. 断点续发：已发送的记录写进 send_state.json，重跑不会重复打扰同一人。
    4. 全部明细写入 send_log_时间戳.csv，方便事后核对。
"""
import argparse
import configparser
import csv
import json
import os
import re
import smtplib
import ssl
import sys
import time
from datetime import datetime
from email.header import Header
from email.mime.text import MIMEText
from email.utils import formataddr, formatdate, make_msgid
from pathlib import Path

BASE = Path(__file__).resolve().parent
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$")


# ---------------------------------------------------------------- 读取名单
def load_rows(excel, sheet, col_email, col_body, start_row, header, col_name=0):
    """返回 [(行号, 邮箱, 姓名, 门票码)]；姓名取名单里的姓名列，没有则用邮箱前缀。"""
    path = Path(excel)
    if not path.is_absolute():
        path = BASE / path
    rows = []
    if path.suffix.lower() in (".csv", ".txt"):
        # utf-8-sig 兼容 Excel 另存的带 BOM 的 CSV；打不开再试 GBK
        last_err = None
        for enc in ("utf-8-sig", "gbk"):
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    data = list(csv.reader(f))
                break
            except UnicodeDecodeError as e:
                last_err = e
        else:
            raise SystemExit(f"[错误] 无法识别 CSV 编码: {last_err}")
    elif path.suffix.lower() in (".xlsx", ".xlsm"):
        try:
            from openpyxl import load_workbook
        except ImportError:
            raise SystemExit("[错误] 读取 xlsx 需要 openpyxl，请先执行: pip install openpyxl")
        wb = load_workbook(path, data_only=True)
        ws = wb[sheet] if sheet else wb.active
        data = [[("" if c is None else str(c)) for c in r] for r in ws.iter_rows(values_only=True)]
    else:
        raise SystemExit(f"[错误] 不支持的名单格式: {path.suffix}，请用 .xlsx 或 .csv")

    # start_row 是第一个数据行；has_header=true 时它上一行是表头，这里直接从 start_row 开始读
    first = max(start_row, 1)
    for idx, row in enumerate(data[first - 1:], start=first):
        if idx < start_row:
            continue
        if idx == start_row and header and start_row == 1:
            continue  # 表头在电子表格第 1 行的情况
        email = (row[col_email - 1] if len(row) >= col_email else "").strip()
        body = (row[col_body - 1] if len(row) >= col_body else "")
        name = (row[col_name - 1] if col_name and len(row) >= col_name else "").strip()
        if not email and not str(body).strip():
            continue  # 整行空白
        rows.append((idx, email, name, body))
    return rows


# ------------------------------------------------------- 正文里的占位符替换
def name_from_email(email):
    """邮箱前缀作为称呼：jiangxin@haiwen-law.com -> jiangxin"""
    return email.split("@", 1)[0].strip() if "@" in email else email.strip()


def render(text, email, index, name="", code=""):
    """支持 {邮箱}{email} {姓名}{name} {门票码}{code} {序号} 占位符，其他内容原样保留。"""
    name = name or name_from_email(email)
    return (text.replace("{邮箱}", email)
                .replace("{email}", email)
                .replace("{姓名}", name)
                .replace("{name}", name)
                .replace("{门票码}", code)
                .replace("{code}", code)
                .replace("{序号}", str(index)))


# ------------------------------------------------------- 一人多码：分组与列表渲染
def group_by_email(rows):
    """把同一邮箱的多行合并成一个收件人。

    返回 [(首行号, 邮箱, 姓名, [码1, 码2, ...])]，保持名单出现顺序；
    码内去重(同一人重复出现的同一个码只保留一次)。
    """
    groups, order = {}, []
    for idx, email, name, body in rows:
        key = email.lower()
        code = str(body).strip()
        if key not in groups:
            groups[key] = (idx, email, name, [])
            order.append(key)
        g = groups[key]
        if not g[2] and str(name).strip():      # 姓名列优先取第一个非空
            groups[key] = (g[0], g[1], name, g[3])
        if code not in groups[key][3]:
            groups[key][3].append(code)
    return [groups[k] for k in order]


def build_body(template, email, idx, name, codes):
    """按码的个数渲染正文。

    单码：{门票码}/{数据} 直接替换，与其他占位符一样处理，输出与旧版完全一致。
    多码：含 {列表} 占位符时渲染在该位置；否则把含 {门票码} 的那一行展开成编号清单，
          后面追加“共 N 个，请核对。”，避免“给你码A码B”糊成一串。
    """
    head = template.replace("{邮箱}", email).replace("{email}", email) \
                   .replace("{姓名}", name or name_from_email(email)) \
                   .replace("{name}", name or name_from_email(email)) \
                   .replace("{序号}", str(idx))

    if len(codes) == 1:
        c = codes[0]
        return (head.replace("{门票码}", c).replace("{数据}", c)
                    .replace("{value}", c).replace("{列表}", c)
                    .replace("{码列表}", c).replace("{codes}", c))

    lines = "\n".join(f"{i}. {c}" for i, c in enumerate(codes, 1))
    total = f"\n\n共 {len(codes)} 个，请核对。"
    for mark in ("{列表}", "{码列表}", "{codes}"):
        if mark in head:
            return head.replace(mark, lines + total)
    # 没写 {列表}：{门票码} 处放第一个码，句子仍通顺；清单接在整段之后
    for mark in ("{门票码}", "{数据}", "{value}"):
        if mark in head:
            return head.replace(mark, codes[0]) + "\n\n清单：\n" + lines + total
    return head.rstrip() + "\n\n清单：\n" + lines + total


def plan_message(template, subject, email, idx, name, codes, mode):
    """生成一封或几封邮件：mode=merge(默认，一人一封) / split(一码一封)。"""
    if mode == "split" and len(codes) > 1:
        out = []
        for i, c in enumerate(codes, 1):
            sub = render(subject, email, idx, name, c)
            out.append((f"{sub}({i}/{len(codes)})", render(template, email, idx, name, c), c))
        return out
    subj = render(subject, email, idx, name, codes[0])
    return [(subj, build_body(template, email, idx, name, codes), "、".join(codes))]


# ------------------------------------------------------------------ 发信
def build_message(cfg, to_addr, subject, body_text, sender_addr):
    msg = MIMEText(body_text, "plain", "utf-8")
    msg["Subject"] = Header(subject, "utf-8").encode()
    msg["From"] = formataddr((str(Header(cfg.get("mail", "from_name", fallback=""), "utf-8")),
                              sender_addr))
    msg["To"] = to_addr
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=sender_addr.split("@")[-1])
    msg["X-Mailer"] = "send_emails.py"
    return msg


def open_smtp(host, port, enc):
    """按加密方式建立连接；STARTTLS 失败时退回明文(Exchange 内网常见)，并打印提示。"""
    last = None
    for attempt in ([enc, "none"] if enc == "starttls" else [enc]):
        try:
            if attempt == "ssl":
                s = smtplib.SMTP_SSL(host, port, timeout=20, context=ssl.create_default_context())
                s.ehlo()
            else:
                s = smtplib.SMTP(host, port, timeout=20)
                s.ehlo()
                if attempt == "starttls":
                    ctx = ssl.create_default_context()
                    ctx.check_hostname = False          # 内网服务器多为自签证书
                    ctx.verify_mode = ssl.CERT_NONE
                    s.starttls(context=ctx)
                    s.ehlo()
            return s
        except Exception as e:  # 端口拒连/证书不对/服务端不支持 STARTTLS
            last = e
            print(f"[连接] {port}/{attempt} 失败: {e}")
    raise SystemExit(f"[错误] 无法连接 {host}:{port} ({enc}) —— {last}\n"
                     f"       尝试顺序：587+starttls → 25+none(内网) → 465+ssl")


def make_sender(cfg):
    """连接 Exchange/内部 SMTP 并登录。

    认证方式 authentication：
      auto   —— 先用 DOMAIN\账号 登录，失败再退回 账号@域 登录（Exchange 最常见两种写法）
      ntlm   —— 强制 DOMAIN\账号
      upn    —— 强制 账号@域
    加密方式 encryption：ssl(465) / starttls(587, 默认) / none(25, 内网明文)
    """
    host = cfg.get("mail", "smtp_host")
    port = int(cfg.get("mail", "smtp_port", fallback="587"))
    enc = cfg.get("mail", "encryption", fallback="starttls").strip().lower()
    user = cfg.get("mail", "user").strip()
    pwd = cfg.get("mail", "password", fallback="").strip()
    if not pwd:
        raise SystemExit("[错误] 请在 config.ini 的 password 填入邮箱密码")

    server = open_smtp(host, port, enc)

    mode = cfg.get("mail", "authentication", fallback="auto").strip().lower()
    aliases = []
    if mode in ("auto", "ntlm"):
        aliases.append(cfg.get("mail", "login_user", fallback=user))
    if mode in ("auto", "upn"):
        aliases.append(user if "@" in user else f"{user}@{cfg.get('mail', 'domain', fallback='')}")

    last_err = None
    for alias in dict.fromkeys(a for a in aliases if a):
        try:
            server.login(alias, pwd)
            print(f"[连接] {host}:{port} 登录成功（账号 {alias}）")
            return server
        except smtplib.SMTPAuthenticationError as e:
            last_err = e
            print(f"[连接] 账号 {alias} 认证失败: {e.smtp_code} {e.smtp_error[:80]!r}")
    raise SystemExit(f"[错误] 所有登录方式均失败: {last_err}\n"
                     f"       请核对 config.ini 的 user/password/domain/login_user，"
                     f"或改用 authentication = ntlm / upn")


def resolve_sender(cfg):
    """发件地址必须是完整邮箱：from_address 优先，否则 账号@域。

    Exchange 会对 MAIL FROM 做地址簿校验，只填 jiangxin 会直接 501 Invalid address。
    """
    addr = cfg.get("mail", "from_address", fallback="").strip()
    if addr and "@" not in addr:
        addr = f"{addr}@{cfg.get('mail', 'domain', fallback='')}"
    if not addr:
        user = cfg.get("mail", "user").strip()
        addr = user if "@" in user else f"{user}@{cfg.get('mail', 'domain', fallback='')}"
    return addr


# ------------------------------------------------------------------ 主流程
def main():
    ap = argparse.ArgumentParser(description="按行个性化发邮件")
    ap.add_argument("--config", default="config.ini")
    ap.add_argument("--excel", help="名单文件(.xlsx/.csv)，默认取 config.ini 中的设置")
    ap.add_argument("--sheet", help="工作表名，默认取 config.ini 中的设置")
    ap.add_argument("--limit", type=int, default=0, help="本轮最多发送 N 条(0=不限)")
    ap.add_argument("--dry-run", action="store_true", default=True, help="预演，不真正发送(默认)")
    ap.add_argument("--no-dry-run", dest="dry_run", action="store_false", help="真的要发信")
    ap.add_argument("--send", action="store_true", help="兼容旧版参数，现在 --no-dry-run 即可")
    ap.add_argument("--retry-failed", action="store_true", help="重发上次失败的收件人")
    ap.add_argument("--split", action="store_true", default=None,
                    help="同一人多个码时拆成多封发(默认合并成一封)")
    ap.add_argument("--merge", dest="split", action="store_false",
                    help="同一人多个码合并成一封(默认)")
    args = ap.parse_args()

    cfg = configparser.ConfigParser(interpolation=None)
    cfg_path = Path(args.config)
    if not cfg_path.is_absolute():
        cfg_path = BASE / cfg_path
    if not cfg_path.exists():
        raise SystemExit(f"[错误] 找不到配置文件 {cfg_path}，请先填写 config.ini")
    cfg.read(cfg_path, encoding="utf-8")

    excel = args.excel or cfg.get("source", "file")
    sheet = args.sheet if args.sheet is not None else cfg.get("source", "sheet", fallback="")
    col_email = cfg.getint("source", "email_column", fallback=1)
    col_body = cfg.getint("source", "body_column", fallback=2)
    start_row = cfg.getint("source", "data_start_row", fallback=2)
    col_name = cfg.getint("source", "name_column", fallback=0)  # 0 = 不读姓名列
    header = cfg.getboolean("source", "has_header", fallback=True)
    subject = cfg.get("mail", "subject")
    template = cfg.get("mail", "body_template", fallback="{门票码}").replace("\\n", "\n")
    interval = cfg.getfloat("mail", "interval", fallback=1.0)
    # merge(默认)：一人多个码合并一封；split：一个码一封
    if args.split is not None:
        mode = "split" if args.split else "merge"
    else:
        mode = cfg.get("mail", "multi_code", fallback="merge").strip().lower() or "merge"
    state_file = BASE / cfg.get("mail", "state_file", fallback="send_state.json")

    rows = load_rows(excel, sheet, col_email, col_body, start_row, header, col_name)
    print(f"[读取] {excel} 共 {len(rows)} 行有效数据" + (f"，工作表: {sheet}" if sheet else ""))

    # 校验（逐行；合格的行再按邮箱分组，一人一封）
    valid, skipped = [], []
    for idx, email, name, body in rows:
        if not EMAIL_RE.match(email):
            skipped.append((idx, email, "邮箱格式不正确")); continue
        if not str(body).strip():
            skipped.append((idx, email, "门票码(B列)为空")); continue
        valid.append((idx, email, name, body))

    people = group_by_email(valid)
    multi = [p for p in people if len(p[3]) > 1]
    dup_codes = len(valid) - sum(len(p[3]) for p in people)

    # 断点续发
    state = {"sent": {}, "failed": {}}
    if state_file.exists():
        state = json.loads(state_file.read_text(encoding="utf-8") or "{}")
        state.setdefault("sent", {}); state.setdefault("failed", {})
    fingerprint = json.dumps([subject, template], ensure_ascii=False)
    if state.get("fingerprint") != fingerprint:
        if state["sent"]:
            print("[提示] 主题或正文模板已变更，已发记录自动作废，本轮会按新模板全部重发。")
        state["sent"] = {}
        state["fingerprint"] = fingerprint
    if args.retry_failed:
        for e in list(state["failed"].keys()):
            state["sent"].pop(e, None)

    todo = [p for p in people if p[1].lower() not in state["sent"]]
    if args.limit:
        todo = todo[: args.limit]
    done = len(people) - len(todo)
    print(f"[统计] 有效 {len(valid)} 行 → {len(people)} 人；跳过 {len(skipped)} 行；"
          f"本轮待发 {len(todo)} 人" + (f"；{done} 人此前已发过(断点记录)" if done else ""))
    if multi:
        print(f"       其中 {len(multi)} 人对应多个码，将合并为一封邮件："
              + "，".join(f"{p[1]}({len(p[3])}个)" for p in multi))
    if dup_codes:
        print(f"       已自动去掉 {dup_codes} 个重复的码")

    if args.dry_run:
        print("\n===== 预演(Dry-Run)：不会真实发送 =====")
        for idx, email, name, codes in todo[:20]:
            src = "名单姓名列" if str(name).strip() else "邮箱前缀"
            print(f"  第{idx}行 -> {email}（{len(codes)} 个码，称呼取自{src}）")
            for sub_j, bod_j, _ in plan_message(template, subject, email, idx, name, codes, mode):
                print(f"      主题: {sub_j}")
                for line in bod_j.split("\n"):
                    print(f"      正文: {line}" if line else "      正文:")
        if len(todo) > 20:
            print(f"  …… 其余 {len(todo) - 20} 条省略")
        skipped_people = [p for p in people if p[1].lower() in state["sent"]]
        for idx, email, name, codes in skipped_people:
            print(f"  [已发] {email}（{'、'.join(codes)}）—— 如需重发请加 --retry-failed")
        for idx, email, reason in skipped:
            print(f"  [跳过] 第{idx}行 {email}: {reason}")
        print(f"\n合并模式: {mode}（merge=一人一封含全部码；split=一个码一封）")
        print("确认无误后执行: python send_emails.py --no-dry-run --limit 3")
        return

    if not todo:
        print("[完成] 没有需要发送的邮件。")
        return

    log_path = BASE / f"send_log_{datetime.now():%Y%m%d_%H%M%S}.csv"
    log_f = open(log_path, "w", encoding="utf-8-sig", newline="")
    log_w = csv.writer(log_f)
    log_w.writerow(["行号", "收件人", "主题", "结果", "说明", "时间"])

    sender_addr = resolve_sender(cfg)
    server = None
    try:
        server = make_sender(cfg)
        print("[连接] 开始发送……")
        for n, (idx, email, name, codes) in enumerate(todo, 1):
            mails = plan_message(template, subject, email, idx, name, codes, mode)
            sent_ok = 0
            try:
                for k, (subject_i, body_i, _) in enumerate(mails, 1):
                    msg = build_message(cfg, email, subject_i, body_i, sender_addr)
                    server.sendmail(sender_addr, [email], msg.as_string())  # 只投这一人
                    sent_ok += 1
                    log_w.writerow([idx, email, subject_i, "成功",
                                    f"{len(codes)}个码" + (f" {k}/{len(mails)}" if len(mails) > 1 else ""),
                                    datetime.now().strftime("%F %T")])
                    if k < len(mails) and interval > 0:
                        time.sleep(interval)  # 同一个人的多封也要间隔，避免被判重发
                state["sent"][email.lower()] = {
                    "row": idx, "codes": codes, "mails": len(mails),
                    "at": datetime.now().isoformat(timespec="seconds")}
                state["failed"].pop(email.lower(), None)
                extra = f"（{len(codes)} 个码）" if len(codes) > 1 else ""
                print(f"  [{n}/{len(todo)}] OK  {email}{extra}")
                log_f.flush()
            except Exception as e:  # 单人失败不影响整体
                if sent_ok:
                    state["sent"][email.lower()] = {
                        "row": idx, "codes": codes[:sent_ok], "mails": sent_ok,
                        "at": datetime.now().isoformat(timespec="seconds")}
                state["failed"][email.lower()] = {"row": idx, "error": str(e)[:200]}
                print(f"  [{n}/{len(todo)}] 失败 {email}: {e}")
                log_w.writerow([idx, email, render(subject, email, idx, name, codes[0]),
                                "失败", str(e)[:200], datetime.now().strftime("%F %T")])
            state_file.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
            log_f.flush()
            if n < len(todo) and interval > 0:
                time.sleep(interval)  # 拉开发送节奏，降低被判垃圾邮件/风控的概率
    except Exception as e:
        print(f"[错误] SMTP 连接/登录失败: {e}")
        print("       检查 config.ini：smtp_host/smtp_port/user/password/domain 是否正确；")
        print("       Exchange 常见组合：587+starttls / 25+none(内网) / 465+ssl；认证可改 authentication=ntlm 或 upn。")
    finally:
        log_f.close()
        try:
            if server:
                server.quit()
        except Exception:
            pass

    print(f"\n[完成] 成功 {len(state['sent'])} 人，失败 {len(state['failed'])} 人")
    print(f"[日志] {log_path}")
    if state["failed"]:
        print("       重跑失败项: python send_emails.py --no-dry-run --send --retry-failed")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[中断] 已发送的记录已保存，重跑会从断点继续。")
        sys.exit(1)
