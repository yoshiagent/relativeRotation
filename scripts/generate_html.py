# -*- coding: utf-8 -*-
"""
從 Excel + CSV 生成每日網頁 (site/index.html)。
"""
from __future__ import annotations
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from pathlib import Path
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook

ROOT = Path(r"C:\CludeHome\projects\relativeRotation")
XLSX = ROOT / "相對輪動模型.xlsx"
CSV  = ROOT / "data" / "prices_daily.csv"
SITE = ROOT / "docs"  # GitHub Pages 支援 / 與 /docs
SITE.mkdir(exist_ok=True)


def read_sheet(name: str) -> pd.DataFrame:
    wb = load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb[name]
    rows = list(ws.values)
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows[1:], columns=rows[0])
    return df


def pair_chart_data(metrics: pd.DataFrame, pair_id: str, n: int = 120):
    """取近 N 日的 Ratio + Z-score 供畫圖。"""
    df = metrics[metrics["配對ID"] == pair_id].tail(n).copy()
    return {
        "labels": [str(d) for d in df["日期"]],
        "ratio":  [float(x) if x is not None else None for x in df["Ratio"]],
        "ratio_ma20": [float(x) if x is not None else None for x in df["Ratio_MA20"]],
        "bb_upper": [float(x) if x is not None else None for x in df["布林上緣2σ"]],
        "bb_lower": [float(x) if x is not None else None for x in df["布林下緣2σ"]],
        "zscore": [float(x) if x is not None else None for x in df["Zscore60"]],
    }


# 各配對的外部詳細頁連結（換股線圖等）。沒設定者保持純文字。
PAIR_LINKS = {
    "P01": "https://yoshiagent.github.io/nan-xin-pairs-trading/%E6%8F%9B%E8%82%A1%E7%B7%9A%E5%9C%96_%E4%BA%92%E5%8B%95%E7%89%88.html",
    "P02": "https://yoshiagent.github.io/liteone-delta-pairs-trading/%E6%8F%9B%E8%82%A1%E7%B7%9A%E5%9C%96_%E4%BA%92%E5%8B%95%E7%89%88.html",
    "P03": "https://yoshiagent.github.io/qihong-shuanghong-pairs-trading/%E6%8F%9B%E8%82%A1%E7%B7%9A%E5%9C%96_%E4%BA%92%E5%8B%95%E7%89%88.html",
    "P04": "https://yoshiagent.github.io/tg-lm-pairs-trading/%E6%8F%9B%E8%82%A1%E7%B7%9A%E5%9C%96_%E4%BA%92%E5%8B%95%E7%89%88.html",
    "P05": "https://yoshiagent.github.io/quanta-wistron-pairs-trading/%E6%8F%9B%E8%82%A1%E7%B7%9A%E5%9C%96_%E4%BA%92%E5%8B%95%E7%89%88.html",
    "P06": "https://yoshiagent.github.io/giant-merida-pairs-trading/%E6%8F%9B%E8%82%A1%E7%B7%9A%E5%9C%96_%E4%BA%92%E5%8B%95%E7%89%88.html",
    "P07": "https://yoshiagent.github.io/shang-da-pairs-trading/%E6%8F%9B%E8%82%A1%E7%B7%9A%E5%9C%96_%E4%BA%92%E5%8B%95%E7%89%88.html",
    "P08": "https://yoshiagent.github.io/nan-hb-pairs-trading/%E6%8F%9B%E8%82%A1%E7%B7%9A%E5%9C%96_%E4%BA%92%E5%8B%95%E7%89%88.html",
}


def phase_color(phase: str) -> str:
    if "過熱" in phase: return "danger"
    if "低估" in phase or "轉強" in phase: return "good"
    if "修正" in phase: return "warn"
    if "末段" in phase: return "warn"
    return "neutral"


def action_color(action: str) -> str:
    s = str(action)
    if "停止" in s: return "danger"
    if "減碼" in s or "降碼" in s: return "warn"
    if "加碼" in s or "換股" in s or "強烈" in s: return "good"
    return "neutral"


def render_card(row, charts) -> str:
    phase = row.get("位置判定", "—") or "—"
    action = row.get("建議動作", "—") or "—"
    star = row.get("⭐補漲標記", "") or ""
    star_html = f'<div class="star">{star}</div>' if star else ""
    pid = row["配對ID"]
    chart_id = f"chart_{pid}"
    link = PAIR_LINKS.get(pid)
    if link:
        pid_html = (f'<a class="pid pid-link" href="{link}" target="_blank" rel="noopener">'
                    f'{pid} <span class="link-icon">↗</span></a>')
    else:
        pid_html = f'<div class="pid">{pid}</div>'
    return f"""
    <div class="card">
        <div class="card-head">
            <div>
                {pid_html}
                <div class="theme">{row.get("主題", "")}</div>
            </div>
            {star_html}
        </div>
        <div class="metrics">
            <div><span class="lbl">Ratio</span><span class="val">{row.get("最新Ratio", "—")}</span></div>
            <div><span class="lbl">Z(60)</span><span class="val z-{('neg' if (row.get('Z60') or 0) < 0 else 'pos')}">{row.get("Z60", "—")}</span></div>
            <div><span class="lbl">Corr(60)</span><span class="val">{row.get("Corr60", "—")}</span></div>
            <div><span class="lbl">A 乖離%</span><span class="val">{row.get("A乖離%", "—")}</span></div>
            <div><span class="lbl">B 乖離%</span><span class="val">{row.get("B乖離%", "—")}</span></div>
            <div><span class="lbl">A Vol</span><span class="val small">{row.get("VolRegime_A", "—")}</span></div>
            <div><span class="lbl">B Vol</span><span class="val small">{row.get("VolRegime_B", "—")}</span></div>
            <div><span class="lbl">ATR A/B</span><span class="val small">{row.get("ATR_A", "—")} / {row.get("ATR_B", "—")}</span></div>
        </div>
        <div class="badges">
            <span class="badge phase-{phase_color(phase)}">{phase}</span>
            <span class="badge action-{action_color(action)}">{action}</span>
        </div>
        <div class="chart-wrap">
            <canvas id="{chart_id}"></canvas>
        </div>
    </div>
    """


def render_table(df: pd.DataFrame, title: str, max_rows: int = 5) -> str:
    if df.empty:
        return f"<h3>{title}</h3><p>無資料</p>"
    # 只取最新一日，每配對一列
    latest_date = df["日期"].max()
    df = df[df["日期"] == latest_date]
    df = df.fillna("—")
    headers = "".join(f"<th>{c}</th>" for c in df.columns)
    body = ""
    for _, row in df.iterrows():
        cells = ""
        for c in df.columns:
            v = row[c]
            cls = ""
            if c == "綜合判定":
                s = str(v)
                if "強烈" in s or "繼續" in s: cls = "td-good"
                elif "適合" in s: cls = "td-highlight"
                elif "觀察" in s or "降碼" in s: cls = "td-warn"
                elif "停止" in s or "不適合" in s: cls = "td-danger"
            elif isinstance(v, bool):
                v = "✓" if v else ""
                cls = "td-good" if v else ""
            cells += f'<td class="{cls}">{v}</td>'
        body += f"<tr>{cells}</tr>"
    return f"""
    <h3>{title} <span class="latest-date">{latest_date}</span></h3>
    <div class="tbl-wrap">
    <table>
        <thead><tr>{headers}</tr></thead>
        <tbody>{body}</tbody>
    </table>
    </div>
    """


def main():
    if not XLSX.exists():
        print(f"[error] 找不到 {XLSX}")
        sys.exit(1)

    dash = read_sheet("儀表板")
    metrics = read_sheet("配對指標")
    entry = read_sheet("進場訊號")
    risk = read_sheet("風險訊號")
    pairs_pool = read_sheet("配對池")

    update_time = datetime.now().strftime("%Y-%m-%d %H:%M")
    data_date = str(dash["更新日期"].iloc[0]) if not dash.empty else "—"

    # 圖表資料
    chart_data = {}
    for _, r in dash.iterrows():
        pid = r["配對ID"]
        chart_data[pid] = pair_chart_data(metrics, pid, n=120)

    # 配對名稱對照（顯示在卡片）
    pair_names = {}
    if not pairs_pool.empty:
        for _, r in pairs_pool.iterrows():
            pair_names[r["配對ID"]] = f"{r['StockA名稱']} / {r['StockB名稱']}"

    # Cards
    cards_html = ""
    for _, r in dash.iterrows():
        pid = r["配對ID"]
        # 加上配對名稱
        r = r.copy()
        r["主題"] = f"{r['主題']} ({pair_names.get(pid,'')})"
        cards_html += render_card(r, chart_data)

    # Tables
    entry_html = render_table(entry, "進場訊號（最新一日）")
    risk_html = render_table(risk, "風險訊號（最新一日）")

    # 統計：補漲、進場、停止
    star_count = sum(1 for v in dash["⭐補漲標記"] if v)
    entry_alert = 0
    if not entry.empty:
        latest = entry[entry["日期"] == entry["日期"].max()]
        entry_alert = sum(1 for v in latest["綜合判定"] if "進場" in str(v))
    stop_count = 0
    if not risk.empty:
        latest_r = risk[risk["日期"] == risk["日期"].max()]
        stop_count = sum(1 for v in latest_r["綜合判定"] if "停止" in str(v))

    html = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex, nofollow">
<title>台股相對輪動儀表板 · {data_date}</title>
<link rel="preconnect" href="https://cdn.jsdelivr.net">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root {{
  --bg: #0d1117; --fg: #e6edf3; --muted: #8b949e;
  --panel: #161b22; --border: #30363d;
  --good: #2ea043; --warn: #d29922; --danger: #f85149;
  --highlight: #f0883e; --accent: #58a6ff;
  --z-neg: #f85149; --z-pos: #2ea043;
}}
* {{ box-sizing: border-box; }}
body {{ margin:0; font-family: -apple-system, "Segoe UI", "Microsoft JhengHei", sans-serif;
  background: var(--bg); color: var(--fg); line-height: 1.5; }}
.container {{ max-width: 1400px; margin: 0 auto; padding: 24px; }}
header {{ display: flex; justify-content: space-between; align-items: baseline;
  border-bottom: 1px solid var(--border); padding-bottom: 16px; margin-bottom: 24px; }}
header h1 {{ margin: 0; font-size: 22px; }}
header .meta {{ color: var(--muted); font-size: 13px; }}
.summary {{ display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap; }}
.sum-card {{ background: var(--panel); border: 1px solid var(--border);
  border-radius: 6px; padding: 12px 20px; min-width: 140px; }}
.sum-card .v {{ font-size: 28px; font-weight: bold; }}
.sum-card .l {{ color: var(--muted); font-size: 13px; }}
.sum-card.star .v {{ color: var(--highlight); }}
.sum-card.alert .v {{ color: var(--accent); }}
.sum-card.stop .v {{ color: var(--danger); }}
.cards {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(380px, 1fr)); gap: 16px; }}
.card {{ background: var(--panel); border: 1px solid var(--border);
  border-radius: 8px; padding: 16px; }}
.card-head {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 12px; }}
.pid {{ font-size: 18px; font-weight: bold; color: var(--accent); }}
.pid-link {{ text-decoration: none; cursor: pointer; display: inline-flex; align-items: center; gap: 4px;
  border-bottom: 1px dashed transparent; transition: border-color 0.15s; }}
.pid-link:hover {{ border-bottom-color: var(--accent); }}
.link-icon {{ font-size: 13px; opacity: 0.65; }}
.theme {{ color: var(--muted); font-size: 13px; }}
.star {{ background: var(--highlight); color: #000; padding: 4px 8px;
  border-radius: 4px; font-size: 12px; font-weight: bold; animation: pulse 2s infinite; }}
@keyframes pulse {{ 0%,100% {{ opacity:1; }} 50% {{ opacity: 0.6; }} }}
.metrics {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 6px 12px; margin-bottom: 10px; }}
.metrics > div {{ display: flex; justify-content: space-between; font-size: 13px; }}
.lbl {{ color: var(--muted); }}
.val {{ font-weight: 600; }}
.val.small {{ font-size: 12px; }}
.val.z-neg {{ color: var(--z-neg); }}
.val.z-pos {{ color: var(--z-pos); }}
.badges {{ display: flex; gap: 8px; flex-wrap: wrap; margin: 10px 0; }}
.badge {{ padding: 4px 10px; border-radius: 4px; font-size: 12px; font-weight: 600; }}
.phase-good, .action-good {{ background: rgba(46,160,67,0.2); color: var(--good); border: 1px solid var(--good); }}
.phase-warn, .action-warn {{ background: rgba(210,153,34,0.2); color: var(--warn); border: 1px solid var(--warn); }}
.phase-danger, .action-danger {{ background: rgba(248,81,73,0.2); color: var(--danger); border: 1px solid var(--danger); }}
.phase-neutral, .action-neutral {{ background: rgba(139,148,158,0.2); color: var(--muted); border: 1px solid var(--border); }}
.chart-wrap {{ height: 140px; margin-top: 8px; }}
section {{ margin-top: 32px; }}
section h2 {{ font-size: 18px; border-left: 4px solid var(--accent); padding-left: 10px; }}
h3 {{ font-size: 15px; color: var(--fg); }}
.latest-date {{ color: var(--muted); font-weight: normal; font-size: 12px; margin-left: 8px; }}
.tbl-wrap {{ overflow-x: auto; border: 1px solid var(--border); border-radius: 6px; }}
table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
thead {{ background: var(--panel); }}
th, td {{ padding: 8px 10px; text-align: center; border-bottom: 1px solid var(--border); }}
th {{ color: var(--muted); font-weight: 600; font-size: 11px; white-space: nowrap; }}
tbody tr:hover {{ background: rgba(88,166,255,0.05); }}
.td-good {{ color: var(--good); font-weight: 600; }}
.td-highlight {{ color: var(--highlight); font-weight: 600; }}
.td-warn {{ color: var(--warn); font-weight: 600; }}
.td-danger {{ color: var(--danger); font-weight: 600; }}
.strategy-doc details {{ background: var(--panel); border: 1px solid var(--border);
  border-radius: 6px; margin-bottom: 10px; padding: 0; }}
.strategy-doc summary {{ cursor: pointer; padding: 12px 16px; font-weight: 600;
  font-size: 14px; color: var(--accent); user-select: none; list-style: none; position: relative; }}
.strategy-doc summary::-webkit-details-marker {{ display: none; }}
.strategy-doc summary::before {{ content: "▶"; display: inline-block; margin-right: 8px;
  font-size: 10px; transition: transform 0.15s; color: var(--muted); }}
.strategy-doc details[open] > summary::before {{ transform: rotate(90deg); }}
.strategy-doc summary:hover {{ background: rgba(88,166,255,0.05); }}
.doc-body {{ padding: 4px 20px 16px 36px; font-size: 13px; line-height: 1.7; }}
.doc-body p {{ margin: 6px 0; }}
.doc-body code {{ background: var(--bg); border: 1px solid var(--border);
  padding: 2px 6px; border-radius: 3px; font-family: "Consolas", monospace; font-size: 12px; }}
.doc-body .quote {{ border-left: 3px solid var(--accent); padding: 4px 12px;
  color: var(--muted); font-style: italic; margin: 12px 0; }}
.doc-body ol.rules, .doc-body ul.rules {{ padding-left: 20px; margin: 8px 0; }}
.doc-body ol.rules li, .doc-body ul.rules li {{ margin: 4px 0; }}
.doc-body .rule-tier {{ width: 100%; margin: 12px 0; border-collapse: collapse;
  border: 1px solid var(--border); border-radius: 4px; overflow: hidden; }}
.doc-body .rule-tier th {{ background: rgba(88,166,255,0.08); color: var(--fg);
  padding: 8px 12px; text-align: left; font-size: 12px; border-bottom: 1px solid var(--border); }}
.doc-body .rule-tier td {{ padding: 6px 12px; border-bottom: 1px solid var(--border); font-size: 13px; }}
.doc-body .rule-tier tr:last-child td {{ border-bottom: none; }}
.star-inline {{ background: var(--highlight); color: #000; padding: 2px 8px;
  border-radius: 3px; font-size: 12px; font-weight: bold; }}
footer {{ margin-top: 48px; padding-top: 16px; border-top: 1px solid var(--border);
  color: var(--muted); font-size: 12px; text-align: center; }}
footer a {{ color: var(--accent); text-decoration: none; }}
footer .local-path {{ margin-top: 8px; }}
footer .local-path code {{ background: var(--panel); border: 1px solid var(--border);
  padding: 2px 8px; border-radius: 4px; color: var(--fg); font-size: 12px;
  font-family: "Consolas", "Courier New", monospace; }}
</style>
</head>
<body>
<div class="container">
<header>
<div>
  <h1>台股雙多頭 Relative Rotation 儀表板</h1>
  <div class="meta">資料日期 <strong>{data_date}</strong> · 最後更新 {update_time}</div>
</div>
<div class="meta">5 組配對 · 60日 Z-score · 滾動 Bollinger</div>
</header>

<div class="summary">
  <div class="sum-card star"><div class="v">{star_count}</div><div class="l">⭐ 補漲候選</div></div>
  <div class="sum-card alert"><div class="v">{entry_alert}</div><div class="l">○ 進場/觀察訊號</div></div>
  <div class="sum-card stop"><div class="v">{stop_count}</div><div class="l">🛑 停止操作</div></div>
  <div class="sum-card"><div class="v">{len(dash)}</div><div class="l">觀察配對</div></div>
</div>

<section>
<h2>📊 配對儀表板</h2>
<div class="cards">{cards_html}</div>
</section>

<section>
<h2>✅ 進場訊號（10 項條件）</h2>
{entry_html}
</section>

<section>
<h2>⚠️ 風險訊號（9 項條件）</h2>
{risk_html}
</section>

<section class="strategy-doc">
<h2>📖 策略邏輯說明</h2>

<details open>
<summary>核心理念</summary>
<div class="doc-body">
<p>本策略屬於 <b>Long-only Pair Rotation</b>（雙多頭相對輪動），非放空、非市場中性、非高頻交易。
透過同產業/同主題的兩檔股票配對，利用 <b>資金輪動</b> 造成的相對低估，
搭配 <b>趨勢、波動率差、均值回歸、Relative Strength</b> 找出「準備補漲」的標的。</p>
<p class="quote">不預測漲跌，而是尋找資金輪動造成的相對低估。</p>
</div>
</details>

<details open>
<summary>配對與 Ratio 定義</summary>
<div class="doc-body">
<p><code>Ratio = StockA 收盤價 / StockB 收盤價</code></p>
<p>Ratio 偏低（Z-score 為負）：A 相對於 B 被低估 → 可能換股至 A。<br>
Ratio 偏高（Z-score 為正）：A 相對於 B 偏貴 → 可能減碼 A、回補 B。</p>
<p>換股以 <b>等資金</b> 而非等股數。例：台達電 500 元一張 ≈ 光寶科 100 元 5 張。</p>
</div>
</details>

<details>
<summary>✅ 進場訊號 — 10 項條件</summary>
<div class="doc-body">
<ol class="rules">
<li>Ratio 的 <b>60日 Z-score &lt; -2</b>（A 相對弱過頭）</li>
<li><b>低檔翻揚</b>：Z-score 上升 + Ratio 站回 5 日均線</li>
<li><b>雙週 MA20 皆向上</b>（A、B 兩檔週線都在多頭趨勢）</li>
<li><b>週 MA20 斜率 &gt; 0</b>（近 4 週持續上升）</li>
<li><b>月 K 線剛脫離長期整理</b>（月收創 12 個月新高）</li>
<li><b>週 K 線 Higher Low</b>（近 4 週低點 &gt; 前 4 週低點）</li>
<li><b>日 K 量縮修正</b>（A 收盤 &lt; MA20 且當日量 &lt; 20 日均量）</li>
<li><b>修正量 &lt; 上漲量</b>（跌日均量 &lt; 漲日均量）</li>
<li><b>20 日均量 &gt; 60 日均量</b>（資金進場）</li>
<li><b>加權指數位於週 MA20 之上</b>（大盤多頭）</li>
</ol>
<table class="rule-tier">
<thead><tr><th>判定等級</th><th>達成條件</th><th>顏色</th></tr></thead>
<tbody>
<tr><td>✓ 強烈進場</td><td>Z &lt; -2 且 達成 ≥ 8 項</td><td class="td-good">綠</td></tr>
<tr><td>○ 適合進場</td><td>Z &lt; -1.5 且 達成 ≥ 6 項</td><td class="td-highlight">橘</td></tr>
<tr><td>△ 觀察</td><td>Z &lt; -1 且 達成 ≥ 4 項</td><td class="td-warn">黃</td></tr>
<tr><td>× 不適合</td><td>Z ≥ 0 或條件不足</td><td class="td-danger">紅</td></tr>
</tbody>
</table>
</div>
</details>

<details>
<summary>⚠️ 風險訊號 — 9 項條件</summary>
<div class="doc-body">
<ol class="rules">
<li><b>Correlation(60) &lt; 0.65</b>（配對相關性下降，警示）</li>
<li><b>Correlation(60) &lt; 0.5</b>（配對失效 → 停止操作）</li>
<li><b>A 股距 MA20 乖離 &gt; 20%</b>（A 過熱）</li>
<li><b>B 股距 MA20 乖離 &gt; 20%</b>（B 過熱）</li>
<li><b>A 週 MA20 轉平/下彎</b>（A 趨勢失效）</li>
<li><b>B 週 MA20 轉平/下彎</b>（B 趨勢失效）</li>
<li><b>加權指數跌破週 MA20</b>（大盤風險）</li>
<li><b>Ratio 突破 120 日布林上緣 + 爆量</b>（結構性脫鉤）</li>
<li><b>時間停損</b>：Z&lt;-2 後 20~30 個交易日仍未回到 -0.5 以內</li>
</ol>
<table class="rule-tier">
<thead><tr><th>判定等級</th><th>觸發條件</th><th>顏色</th></tr></thead>
<tbody>
<tr><td>🛑 停止操作</td><td>觸發 ②、⑦ 或 ⑧ 任一項</td><td class="td-danger">紅</td></tr>
<tr><td>⚠ 降碼</td><td>其餘條件中觸發 ≥ 2 項</td><td class="td-warn">黃</td></tr>
<tr><td>✓ 繼續</td><td>觸發 &lt; 2 項</td><td class="td-good">綠</td></tr>
</tbody>
</table>
</div>
</details>

<details>
<summary>📊 位置判定 — 7 階段</summary>
<div class="doc-body">
<table class="rule-tier">
<thead><tr><th>Z-score 區間</th><th>位置</th><th>意義</th></tr></thead>
<tbody>
<tr><td>任一股 MA20 乖離 &gt; 20%</td><td>過熱段</td><td>避免追價</td></tr>
<tr><td>Z &lt; -2</td><td>修正段（低估）</td><td>準備進場觀察</td></tr>
<tr><td>-2 ≤ Z &lt; -1</td><td>重新轉強</td><td>補漲訊號區</td></tr>
<tr><td>-1 ≤ Z &lt; 0</td><td>主升初段</td><td>趨勢確立</td></tr>
<tr><td>0 ≤ Z &lt; 1</td><td>主升中段</td><td>持有續抱</td></tr>
<tr><td>1 ≤ Z &lt; 2</td><td>主升末段</td><td>準備減碼</td></tr>
<tr><td>Z ≥ 2</td><td>過熱段（強勢）</td><td>減碼回補</td></tr>
</tbody>
</table>
</div>
</details>

<details>
<summary>⭐ 補漲標記觸發邏輯</summary>
<div class="doc-body">
<p>同時滿足下列三項時，配對卡片右上角會出現 <span class="star-inline">★ 即將補漲</span>：</p>
<ol class="rules">
<li><b>Z-score &lt; -1.5</b>（明確低估）</li>
<li>進場訊號達 <b>觀察或進場</b> 等級以上</li>
<li><b>無停止操作</b>風險訊號</li>
</ol>
<p>此標記代表「最像即將開始補漲」的候選，建議搭配當日成交量、新聞面再做最終判斷。</p>
</div>
</details>

<details>
<summary>🎯 交易哲學與重要原則</summary>
<div class="doc-body">
<ul class="rules">
<li>優先 <b>趨勢向上</b> 的股票（雙多頭環境）</li>
<li>避免 <b>空頭中的均值回歸</b>（接刀風險）</li>
<li>優先 <b>月線剛突破整理</b> 的標的</li>
<li>優先 <b>週線 Higher Low</b> 結構</li>
<li>優先 <b>日線量縮修正</b>（健康回測）</li>
<li>避免 <b>爆量情緒末升段</b>（追高風險）</li>
<li>避免 <b>新聞極度熱門後追價</b></li>
<li>重視 <b>成交量與資金流向</b></li>
<li>重視 <b>法人型慢牛股</b></li>
<li>核心是 <b>Relative Rotation</b>，而非猜大盤</li>
</ul>
</div>
</details>

</section>

<footer>
<p>策略：Long-only Pair Rotation · 不放空、不對沖、不預測大盤</p>
<p>本網頁僅為個人研究紀錄，不構成投資建議 · <a href="https://github.com/yoshiagent/relativeRotation">原始碼</a></p>
<p class="local-path">本機資料夾：<code>C:\\CludeHome\\projects\\relativeRotation</code></p>
</footer>
</div>

<script>
const CHARTS = {json.dumps(chart_data, ensure_ascii=False)};
const COMMON_OPTS = {{
  responsive: true, maintainAspectRatio: false,
  plugins: {{ legend: {{ display: true, labels: {{ color: '#8b949e', font: {{ size: 10 }} }} }},
    tooltip: {{ mode: 'index', intersect: false }} }},
  scales: {{
    x: {{ display: false }},
    y: {{ ticks: {{ color: '#8b949e', font: {{ size: 10 }} }}, grid: {{ color: 'rgba(139,148,158,0.1)' }} }}
  }},
  interaction: {{ mode: 'index', intersect: false }},
  elements: {{ point: {{ radius: 0 }}, line: {{ borderWidth: 1.5 }} }}
}};
for (const pid in CHARTS) {{
  const d = CHARTS[pid];
  const ctx = document.getElementById('chart_' + pid);
  if (!ctx) continue;
  new Chart(ctx, {{
    type: 'line',
    data: {{
      labels: d.labels,
      datasets: [
        {{ label: 'Ratio', data: d.ratio, borderColor: '#58a6ff', backgroundColor: 'rgba(88,166,255,0.1)', fill: false }},
        {{ label: 'MA20', data: d.ratio_ma20, borderColor: '#8b949e', borderDash: [4,4], fill: false }},
        {{ label: '布林上', data: d.bb_upper, borderColor: 'rgba(248,81,73,0.4)', fill: false }},
        {{ label: '布林下', data: d.bb_lower, borderColor: 'rgba(46,160,67,0.4)', fill: '+1', backgroundColor: 'rgba(46,160,67,0.05)' }}
      ]
    }},
    options: COMMON_OPTS
  }});
}}
</script>
</body>
</html>
"""
    out = SITE / "index.html"
    out.write_text(html, encoding="utf-8")

    # robots.txt
    (SITE / "robots.txt").write_text("User-agent: *\nDisallow: /\n", encoding="utf-8")

    # 同時把當前信號狀態存成 JSON（給 email 對比用）
    sig_state = []
    for _, r in dash.iterrows():
        sig_state.append({
            "配對ID": r["配對ID"],
            "位置判定": r.get("位置判定", ""),
            "建議動作": str(r.get("建議動作", "")),
            "補漲標記": str(r.get("⭐補漲標記", "")),
        })
    # 進場訊號最新一日
    if not entry.empty:
        latest = entry[entry["日期"] == entry["日期"].max()]
        for _, r in latest.iterrows():
            for s in sig_state:
                if s["配對ID"] == r["配對ID"]:
                    s["進場判定"] = str(r.get("綜合判定", ""))
    (ROOT / "data" / "signal_state_today.json").write_text(
        json.dumps({"date": data_date, "pairs": sig_state}, ensure_ascii=False, indent=2),
        encoding="utf-8")

    print(f"[ok] 寫入 {out}")
    print(f"[ok] 寫入 {SITE / 'robots.txt'}")
    print(f"[stat] 補漲={star_count}  進場/觀察={entry_alert}  停止={stop_count}")


if __name__ == "__main__":
    main()
