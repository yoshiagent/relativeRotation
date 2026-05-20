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
footer {{ margin-top: 48px; padding-top: 16px; border-top: 1px solid var(--border);
  color: var(--muted); font-size: 12px; text-align: center; }}
footer a {{ color: var(--accent); text-decoration: none; }}
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

<footer>
<p>策略：Long-only Pair Rotation · 不放空、不對沖、不預測大盤</p>
<p>本網頁僅為個人研究紀錄，不構成投資建議 · <a href="https://github.com/yoshiagent/relativeRotation">原始碼</a></p>
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
