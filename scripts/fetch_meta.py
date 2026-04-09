"""
fetch_meta.py
從 Meta Marketing API 拉廣告報表，輸出 CSV（格式與 analyze_ads.py 相容）

用法：
  python scripts/fetch_meta.py --period full   # 本週一～今天（週五執行）或上週（週一執行）
  python scripts/fetch_meta.py --since 2026-03-31 --until 2026-04-06  # 自訂日期

輸出（預設路徑）：
  /tmp/meta_full.csv
  /tmp/meta_7d.csv
  /tmp/meta_4d.csv
"""

import warnings; warnings.filterwarnings("ignore")
import csv, json, os, sys, argparse
from datetime import date, timedelta
import requests
from dotenv import load_dotenv

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), "..", ".env"))

TOKEN    = os.getenv("META_ACCESS_TOKEN")
ACCOUNT  = os.getenv("META_AD_ACCOUNT_ID")   # e.g. act_552610531057805

if not TOKEN or not ACCOUNT:
    print("ERROR: META_ACCESS_TOKEN 或 META_AD_ACCOUNT_ID 未設定", file=sys.stderr)
    sys.exit(1)

BASE = "https://graph.facebook.com/v19.0"

# ── 日期計算 ──────────────────────────────────────────────────────────────────

def this_monday():
    today = date.today()
    return today - timedelta(days=today.weekday())

def last_monday():
    return this_monday() - timedelta(weeks=1)

def last_sunday():
    return this_monday() - timedelta(days=1)

def auto_full_range():
    """週一執行：上週一～上週日；其他時間：本週一～昨天"""
    today = date.today()
    if today.weekday() == 0:   # 週一
        return last_monday(), last_sunday()
    else:
        return this_monday(), today - timedelta(days=1)

# ── Meta API 查詢 ─────────────────────────────────────────────────────────────

FIELDS = ",".join([
    "ad_name", "ad_id", "adset_name",
    "spend", "impressions", "reach",
    "actions", "cost_per_action_type",
])

def fetch_insights(since: date, until: date) -> list:
    params = {
        "access_token": TOKEN,
        "level": "ad",
        "fields": FIELDS,
        "time_range": json.dumps({"since": str(since), "until": str(until)}),
        "limit": 500,
    }
    rows = []
    url = f"{BASE}/{ACCOUNT}/insights"
    while url:
        r = requests.get(url, params=params, timeout=30)
        r.raise_for_status()
        data = r.json()
        rows.extend(data.get("data", []))
        url = data.get("paging", {}).get("next")
        params = {}   # paging URL already contains all params
    return rows

def extract_leads(row: dict) -> tuple:
    """從 actions / cost_per_action_type 提取名單數與 CPL"""
    leads = 0
    cpl   = None
    for a in row.get("actions", []):
        if a.get("action_type") == "lead":
            leads = int(float(a.get("value", 0)))
    for a in row.get("cost_per_action_type", []):
        if a.get("action_type") == "lead":
            cpl = float(a.get("value", 0))
    return leads, cpl

def to_csv_rows(raw: list) -> list:
    """轉換成 analyze_ads.py 期望的中文欄位 CSV 格式"""
    out = []
    for r in raw:
        leads, cpl = extract_leads(r)
        out.append({
            "廣告名稱":          r.get("ad_name", ""),
            "廣告編號":          r.get("ad_id", ""),
            "廣告組合名稱":      r.get("adset_name", ""),
            "廣告投遞":          "",          # insights API 不回傳投遞狀態，留空
            "成果":              leads,
            "每次成果成本":      round(cpl, 4) if cpl else "",
            "花費金額 (USD)":    round(float(r.get("spend", 0)), 2),
            "觸及人數":          r.get("reach", 0),
            "曝光次數":          r.get("impressions", 0),
        })
    return out

def save_csv(rows: list, path: str):
    if not rows:
        print(f"[警告] 無資料，跳過儲存：{path}", file=sys.stderr)
        return
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)
    print(f"已儲存 {len(rows)} 支廣告 → {path}")

# ── CLI ───────────────────────────────────────────────────────────────────────

p = argparse.ArgumentParser()
p.add_argument("--since",        help="完整期間起始日 YYYY-MM-DD")
p.add_argument("--until",        help="完整期間結束日 YYYY-MM-DD")
p.add_argument("--out-full",     default="/tmp/meta_full.csv")
p.add_argument("--out-7d",       default="/tmp/meta_7d.csv")
p.add_argument("--out-4d",       default="/tmp/meta_4d.csv")
args = p.parse_args()

today = date.today()

if args.since and args.until:
    full_since = date.fromisoformat(args.since)
    full_until = date.fromisoformat(args.until)
else:
    full_since, full_until = auto_full_range()

d7_since = today - timedelta(days=6)
d4_since = today - timedelta(days=3)

print(f"完整期間：{full_since} ～ {full_until}")
print(f"近 7 天：{d7_since} ～ {today}")
print(f"近 4 天：{d4_since} ～ {today}")

print("\n拉取完整期間報表...")
save_csv(to_csv_rows(fetch_insights(full_since, full_until)), args.out_full)

print("拉取近 7 天報表...")
save_csv(to_csv_rows(fetch_insights(d7_since, today)), args.out_7d)

print("拉取近 4 天報表...")
save_csv(to_csv_rows(fetch_insights(d4_since, today)), args.out_4d)

print("\n完成。")
