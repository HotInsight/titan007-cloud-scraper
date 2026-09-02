from __future__ import annotations

import argparse
import os
import smtplib
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo


def required(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"GitHub Secret {name} 未设置")
    return value


def send(message: EmailMessage) -> None:
    username = required("GMAIL_USERNAME")
    password = required("GMAIL_APP_PASSWORD").replace(" ", "")
    message["From"] = username
    message["To"] = required("EMAIL_TO")
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
        smtp.login(username, password)
        smtp.send_message(message)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", type=Path)
    parser.add_argument("--failure-log", type=Path)
    args = parser.parse_args()
    now = datetime.now(ZoneInfo("Asia/Bangkok"))
    msg = EmailMessage()

    if args.file:
        if not args.file.exists():
            raise FileNotFoundError(args.file)
        msg["Subject"] = f"新球体育单场数据 {now:%Y-%m-%d}"
        msg.set_content(f"已完成泰国时间 {now:%Y-%m-%d %H:%M} 的自动抓取，Excel见附件：{args.file.name}")
        msg.add_attachment(args.file.read_bytes(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=args.file.name)
    else:
        msg["Subject"] = f"[失败] 新球体育自动抓取 {now:%Y-%m-%d}"
        body = "今天的自动抓取失败。请登录GitHub查看Actions运行记录和诊断附件。"
        if args.failure_log and args.failure_log.exists():
            log_text = args.failure_log.read_text(encoding="utf-8", errors="replace")[-12000:]
            body += "\n\n日志末尾：\n" + log_text
            msg.add_attachment(args.failure_log.read_bytes(), maintype="text", subtype="plain", filename="scraper.log")
        msg.set_content(body)
    send(msg)


if __name__ == "__main__":
    main()
