"""
pipeline.py — 兩層廣告分析流程協調腳本

Layer 1（廣告成效）：
  Meta + Systeme 本期新名單（order:desc，遇舊日期停止）

Layer 2（轉換追蹤）：
  Systeme 全量 + Notion 諮詢資料

用法：
  python scripts/pipeline.py --since 2026-03-31 --until 2026-04-06
      --notion-raw output/raw_notion.json --output output/report.xlsx
      [--period "2026年4月W14"]
      [--skip-layer2-fetch]   # 若 output/systeme_leads.csv 已存在，跳過全量抓取
      [--skip-meta-fetch]     # 使用已快取的 Meta 資料

完整跑法（自動決定日期）：
  python scripts/pipeline.py --notion-raw output/raw_notion.json
"""

import argparse
import json
import os
import subprocess
import sys
from datetime import date, timedelta

SCRIPTS = os.path.dirname(os.path.abspath(__file__))
ROOT    = os.path.dirname(SCRIPTS)
OUT     = os.path.join(ROOT, "output")
os.makedirs(OUT, exist_ok=True)


# ── 工具函式 ──────────────────────────────────────────────────────────────────

def run(cmd: list, label: str) -> int:
    print(f"\n{'='*60}")
    print(f"▶ {label}")
    print(f"  {' '.join(cmd)}")
    print('='*60)
    result = subprocess.run(cmd, cwd=ROOT)
    if result.returncode != 0:
        print(f"[ERROR] {label} 失敗（exit code {result.returncode}）", file=sys.stderr)
    return result.returncode


def auto_period():
    """週一執行：上週一～上週日；其他時間：本週一～昨天"""
    today = date.today()
    weekday = today.weekday()
    if weekday == 0:
        since = today - timedelta(weeks=1)
        until = today - timedelta(days=1)
    else:
        since = today - timedelta(days=weekday)
        until = today - timedelta(days=1)
    return since, until


# ── CLI ───────────────────────────────────────────────────────────────────────

p = argparse.ArgumentParser(description="兩層廣告分析流程")
p.add_argument("--since",  default=None, help="期間起始日 YYYY-MM-DD（未填自動計算）")
p.add_argument("--until",  default=None, help="期間結束日 YYYY-MM-DD（未填自動計算）")
p.add_argument("--notion-raw", default=None,
               help="Notion MCP 原始 JSON 路徑（省略則跳過 Notion 層）")
p.add_argument("--output", default=os.path.join(OUT, "ads_report.xlsx"))
p.add_argument("--period", default="", help="報表顯示用期間標籤，e.g. 2026年4月W14")
p.add_argument("--skip-layer2-fetch", action="store_true",
               help="跳過 Systeme 全量抓取（使用 output/systeme_leads.csv 快取）")
p.add_argument("--skip-meta-fetch", action="store_true",
               help="跳過 Meta 抓取（使用快取的 output/meta_*.csv）")
p.add_argument("--layer2-since", default="2026-03-01",
               help="Layer 2 歷史名單起始日（預設 2026-03-01）")
p.add_argument("--cpl-threshold", type=float, default=None)
args = p.parse_args()

# 決定期間
if args.since and args.until:
    since = date.fromisoformat(args.since)
    until = date.fromisoformat(args.until)
else:
    since, until = auto_period()
    print(f"自動期間：{since} ～ {until}")

period_label = args.period or f"{since.strftime('%Y年%-m月')} {since}～{until}"

errors = []

# ════════════════════════════════════════════════════════════════
# Layer 1：Meta + Systeme 本期新名單
# ════════════════════════════════════════════════════════════════
print("\n" + "█"*60)
print("█  Layer 1：廣告成效（Meta + Systeme 本期新名單）")
print("█"*60)

# 1a. Meta 廣告報表
if not args.skip_meta_fetch:
    rc = run(
        [sys.executable, os.path.join(SCRIPTS, "fetch_meta.py"),
         "--since", str(since), "--until", str(until)],
        "Meta 廣告報表抓取"
    )
    if rc != 0:
        errors.append("Meta fetch 失敗")
else:
    print("⏭  跳過 Meta 抓取，使用快取。")

# 1b. Systeme 本期新名單（order:desc，遇舊日期停止）
rc = run(
    [sys.executable, os.path.join(SCRIPTS, "fetch_systeme.py"),
     "--mode", "new", "--since", str(since)],
    f"Systeme 本期新名單（since={since}）"
)
if rc != 0:
    errors.append("Systeme Layer 1 fetch 失敗")

# ════════════════════════════════════════════════════════════════
# Layer 2：Systeme 全量 + Notion（轉換追蹤）
# ════════════════════════════════════════════════════════════════
print("\n" + "█"*60)
print("█  Layer 2：轉換追蹤（Systeme 全量 + Notion）")
print("█"*60)

# 2a. Systeme 全量（有快取且 < 24 小時則跳過）
LAYER2_CACHE = os.path.join(OUT, "systeme_leads.csv")
CACHE_MAX_AGE_HOURS = 24

def cache_is_fresh(path: str, max_hours: int) -> bool:
    if not os.path.exists(path):
        return False
    import time
    age_hours = (time.time() - os.path.getmtime(path)) / 3600
    return age_hours < max_hours

if args.skip_layer2_fetch:
    print("⏭  --skip-layer2-fetch：跳過全量抓取。")
elif cache_is_fresh(LAYER2_CACHE, CACHE_MAX_AGE_HOURS):
    mtime = __import__('datetime').datetime.fromtimestamp(os.path.getmtime(LAYER2_CACHE))
    print(f"⏭  快取仍有效（{mtime:%Y-%m-%d %H:%M}），跳過全量抓取。")
else:
    layer2_cmd = [sys.executable, os.path.join(SCRIPTS, "fetch_systeme.py"),
                  "--mode", "new",
                  "--since", args.layer2_since,
                  "--out", os.path.join(OUT, "systeme_leads.json")]
    rc = run(layer2_cmd, f"Systeme 歷史名單抓取（since={args.layer2_since}）")
    if rc != 0:
        errors.append("Systeme Layer 2 全量 fetch 失敗")

# 2b. Notion 諮詢資料
notion_json = None
if args.notion_raw:
    notion_json = os.path.join(OUT, "notion_consults.json")
    rc = run(
        [sys.executable, os.path.join(SCRIPTS, "fetch_notion.py"),
         "--raw", args.notion_raw, "--output", notion_json],
        "Notion 諮詢資料清洗"
    )
    if rc != 0:
        errors.append("Notion 資料清洗失敗")
        notion_json = None
else:
    print("⏭  未提供 --notion-raw，跳過 Notion 層。")

# ════════════════════════════════════════════════════════════════
# 組合：生成報表
# ════════════════════════════════════════════════════════════════
print("\n" + "█"*60)
print("█  生成 Excel 報表")
print("█"*60)

# 決定 Layer 2 全量名單路徑
all_leads_csv = os.path.join(OUT, "systeme_leads.csv")
if not os.path.exists(all_leads_csv):
    all_leads_csv = os.path.join(OUT, "systeme_new_leads.csv")
    print(f"[WARN] 全量名單不存在，改用本期名單：{all_leads_csv}")

analyze_cmd = [
    sys.executable, os.path.join(SCRIPTS, "analyze_ads.py"),
    "--full",   os.path.join(OUT, "meta_full.csv"),
    "--7d",     os.path.join(OUT, "meta_7d.csv"),
    "--4d",     os.path.join(OUT, "meta_4d.csv"),
    "--leads",  os.path.join(OUT, "systeme_new_leads.csv"),   # Layer 1：本期新名單
    "--all-leads", all_leads_csv,               # Layer 2：全量名單
    "--output", args.output,
    "--period", period_label,
]
if notion_json:
    analyze_cmd += ["--notion", notion_json]
if args.cpl_threshold:
    analyze_cmd += ["--cpl-threshold", str(args.cpl_threshold)]

rc = run(analyze_cmd, "analyze_ads.py 生成報表")
if rc != 0:
    errors.append("analyze_ads 失敗")

# ════════════════════════════════════════════════════════════════
# 最終摘要
# ════════════════════════════════════════════════════════════════
print("\n" + "═"*60)
print("最終摘要")
print("═"*60)
summary = {
    "period": f"{since} ～ {until}",
    "period_label": period_label,
    "output": args.output,
    "layer1": {
        "meta_csv": os.path.join(OUT, "meta_full.csv"),
        "systeme_new_leads": os.path.join(OUT, "systeme_new_leads.csv"),
    },
    "layer2": {
        "systeme_all_leads": all_leads_csv,
        "notion_consults": notion_json,
    },
    "errors": errors,
    "status": "SUCCESS" if not errors else f"PARTIAL（{len(errors)} 項錯誤）",
}
print(json.dumps(summary, ensure_ascii=False, indent=2))

if errors:
    sys.exit(1)
