# -*- coding: utf-8 -*-
"""
比對今日 vs 昨日信號狀態，有變化時寄信通知。

觸發條件：
  ① 出現新 ⭐補漲標記
  ② 進場訊號升級（不適合→觀察→適合→強烈）
  ③ 出現 🛑 停止操作（從非停止 → 停止）
  ④ 位置判定變動（主升→過熱、修正→重新轉強 等）

設定：讀取 .env 檔案
  GMAIL_USER=hwangprobot@gmail.com
  GMAIL_APP_PASSWORD=xxxxxxxxxxxxxxxx
  MAIL_TO=hwangprobot@gmail.com
"""
from __future__ import annotations
import sys, io, os, json, smtplib, ssl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from pathlib import Path
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr, formatdate

ROOT = Path(r"C:\CludeHome\projects\relativeRotation")
TODAY_FILE  = ROOT / "data" / "signal_state_today.json"
YDAY_FILE   = ROOT / "data" / "signal_state_yesterday.json"
ENV_FILE    = ROOT / ".env"
SITE_URL    = "https://yoshiagent.github.io/relativeRotation/"


# 進場判定等級（越高越強）
ENTRY_RANK = {
    "× 不適合": 0, "× 不適合（Z≥0）": 0,
    "△ 觀察": 1,
    "○ 適合進場": 2,
    "✓ 強烈進場": 3,
}


def load_env():
    if not ENV_FILE.exists():
        return {}
    env = {}
    for line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"): continue
        if "=" in line:
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip().strip('"').strip("'")
    return env


def diff_signals(today: dict, yday: dict) -> list:
    """回傳變化清單，每項為 (配對ID, 變化類型, 詳細描述)。"""
    if not yday:
        return []  # 第一天沒得比，不寄信
    y_map = {p["配對ID"]: p for p in yday.get("pairs", [])}
    changes = []
    for p in today.get("pairs", []):
        pid = p["配對ID"]
        y = y_map.get(pid, {})

        # ① 新補漲標記
        if p.get("補漲標記") and not y.get("補漲標記"):
            changes.append((pid, "star", f"出現新 ⭐ 補漲標記"))

        # ② 進場訊號升級
        t_lvl = ENTRY_RANK.get(p.get("進場判定", ""), 0)
        y_lvl = ENTRY_RANK.get(y.get("進場判定", ""), 0)
        if t_lvl > y_lvl:
            changes.append((pid, "entry_up",
                f"進場訊號升級：{y.get('進場判定','—')} → {p.get('進場判定','—')}"))

        # ③ 出現停止操作
        t_act = str(p.get("建議動作", ""))
        y_act = str(y.get("建議動作", ""))
        if "停止" in t_act and "停止" not in y_act:
            changes.append((pid, "stop", f"出現 🛑 停止操作"))

        # ④ 位置判定變動
        t_phase = p.get("位置判定", "")
        y_phase = y.get("位置判定", "")
        if t_phase and y_phase and t_phase != y_phase:
            changes.append((pid, "phase", f"位置判定：{y_phase} → {t_phase}"))

    return changes


def render_mail_body(changes: list, today: dict) -> str:
    date = today.get("date", "—")
    pairs = {p["配對ID"]: p for p in today.get("pairs", [])}

    rows = ""
    for pid, kind, desc in changes:
        p = pairs.get(pid, {})
        icon = {"star": "⭐", "entry_up": "🟢", "stop": "🛑", "phase": "🔄"}.get(kind, "•")
        rows += f"""
        <tr>
          <td>{icon}</td>
          <td><b>{pid}</b></td>
          <td>{desc}</td>
          <td>{p.get("位置判定","")}</td>
          <td>{p.get("建議動作","")}</td>
        </tr>"""

    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:-apple-system,'Microsoft JhengHei',sans-serif; color:#24292f;">
<div style="max-width:680px; margin:0 auto; padding:24px;">
<h2 style="color:#0969da; border-bottom:2px solid #0969da; padding-bottom:8px;">📈 相對輪動模型 · 信號異動</h2>
<p>資料日期：<b>{date}</b> · 共 <b>{len(changes)}</b> 個變化</p>

<table style="width:100%; border-collapse:collapse; margin:16px 0; font-size:14px;">
  <thead style="background:#f6f8fa;">
    <tr>
      <th style="padding:8px;border:1px solid #d0d7de;"></th>
      <th style="padding:8px;border:1px solid #d0d7de;">配對</th>
      <th style="padding:8px;border:1px solid #d0d7de;">變化</th>
      <th style="padding:8px;border:1px solid #d0d7de;">位置</th>
      <th style="padding:8px;border:1px solid #d0d7de;">建議</th>
    </tr>
  </thead>
  <tbody style="background:white;">{rows}
  </tbody>
</table>

<p style="margin-top:24px;">
  <a href="{SITE_URL}" style="display:inline-block; background:#0969da; color:white;
     padding:10px 20px; text-decoration:none; border-radius:6px;">查看完整儀表板 →</a>
</p>

<hr style="margin:24px 0; border:none; border-top:1px solid #d0d7de;">
<p style="color:#656d76; font-size:12px;">
  此信由 relativeRotation 自動發送 · {datetime.now():%Y-%m-%d %H:%M:%S}<br>
  本系統僅為個人研究紀錄，不構成投資建議。
</p>
</div></body></html>
"""


def send_mail(subject: str, html_body: str, env: dict) -> bool:
    user = env.get("GMAIL_USER")
    pw   = env.get("GMAIL_APP_PASSWORD")
    to   = env.get("MAIL_TO", user)
    if not user or not pw:
        print("[skip] 缺少 GMAIL_USER / GMAIL_APP_PASSWORD")
        return False
    if pw in ("YOUR_APP_PASSWORD_HERE", ""):
        print("[skip] 尚未設定 GMAIL_APP_PASSWORD")
        return False

    msg = MIMEMultipart("alternative")
    msg["From"] = formataddr(("台股相對輪動模型", user))
    msg["To"] = to
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ctx) as s:
        s.login(user, pw)
        s.sendmail(user, [to], msg.as_string())
    print(f"[ok] 已寄信至 {to}")
    return True


def main():
    if not TODAY_FILE.exists():
        print("[skip] 尚無今日信號狀態，請先執行 generate_html.py")
        return

    today = json.loads(TODAY_FILE.read_text(encoding="utf-8"))
    yday = json.loads(YDAY_FILE.read_text(encoding="utf-8")) if YDAY_FILE.exists() else {}

    # 若資料日期相同（同日多次跑），不比也不寄
    if yday and today.get("date") == yday.get("date"):
        # 同日重跑，仍滾動以 today 蓋 yday（保留最新基準）
        # 但不寄信，避免重複通知
        print(f"[skip] 同日重跑（{today.get('date')}），不寄信")
        # 不滾動，保留原 yday
        return

    changes = diff_signals(today, yday)
    if not changes:
        print(f"[info] 無信號變化（{today.get('date')}），不寄信")
    else:
        env = load_env()
        date = today.get("date", "—")
        subject = f"[相對輪動] {date} · {len(changes)} 個信號變化"
        # 摘要 subject
        kinds = set(c[1] for c in changes)
        kind_label = {"star":"⭐補漲", "entry_up":"進場升級", "stop":"🛑停止", "phase":"位置變動"}
        tag = " / ".join(kind_label[k] for k in kinds if k in kind_label)
        if tag:
            subject = f"[相對輪動] {date} · {tag}"
        body = render_mail_body(changes, today)
        try:
            send_mail(subject, body, env)
        except Exception as e:
            print(f"[error] 寄信失敗: {e}")

    # 滾動：今日 → 昨日
    YDAY_FILE.write_text(TODAY_FILE.read_text(encoding="utf-8"), encoding="utf-8")
    print(f"[ok] 信號狀態已滾動")


if __name__ == "__main__":
    main()
