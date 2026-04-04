"""
BRAND 廣告分析 — Notion 諮詢資料取得輔助腳本

使用方式（在 Claude 對話中執行此腳本前，需先透過 Notion MCP 取得資料）：

步驟：
1. Claude 使用 notion-query-data-sources 工具查詢資料庫 346b8e54-61be-49a4-8379-2a1f6a33075b
2. 將回傳的原始 JSON 儲存至 raw_notion.json
3. 執行此腳本清洗、過濾並轉換為 analyze_ads.py 所需格式

用法：
  python fetch_notion.py --raw raw_notion.json --output notion_consults.json

過濾規則：
- 排除 客戶名稱 包含「一人事業」的記錄（不同專案）
- 排除 狀態 為「未出席」或「取消諮詢」的記錄（不算有效諮詢）
"""
import json, argparse, sys
from datetime import datetime

p = argparse.ArgumentParser()
p.add_argument('--raw',    required=True, help='Notion MCP 回傳的原始 JSON 路徑')
p.add_argument('--output', required=True, help='清洗後輸出 JSON 路徑')
p.add_argument('--include-cancelled', action='store_true',
               help='若加上此旗標，則保留取消/未出席的記錄（預設排除）')
args = p.parse_args()

# ── Load raw Notion data ────────────────────────────────────────────────────
with open(args.raw, encoding='utf-8') as f:
    raw = json.load(f)

# Raw data could be a list directly or wrapped in a dict
if isinstance(raw, dict):
    # Try common wrapper keys
    records = (raw.get('results') or raw.get('pages') or
               raw.get('data') or raw.get('rows') or [])
else:
    records = raw

# ── Status categories ───────────────────────────────────────────────────────
EXCLUDE_STATUS = {'未出席', '取消諮詢'}

def extract_field(record, *keys):
    """Try multiple key names, return first match."""
    for k in keys:
        if k in record and record[k] is not None and record[k] != '':
            return record[k]
    return None

def normalize_email(v):
    if not v: return ''
    if isinstance(v, list): v = v[0] if v else ''
    return str(v).strip().lower()

def normalize_status(v):
    if not v: return ''
    if isinstance(v, list): v = ', '.join(str(x) for x in v)
    return str(v).strip()

def normalize_date(v):
    if not v: return ''
    if isinstance(v, dict):
        v = v.get('start') or v.get('date') or ''
    return str(v).strip()

# ── Process records ─────────────────────────────────────────────────────────
output = []
skipped_iys  = 0  # 一人事業
skipped_stat = 0  # 無效狀態

for r in records:
    # Support both flat dict and nested Notion page format
    props = r.get('properties', r)

    name    = extract_field(props, '客戶名稱', 'name', '姓名', 'Name')
    email   = extract_field(props, 'Email', 'email', '電子郵件', 'E-mail')
    status  = extract_field(props, '狀態', 'status', 'Status')
    date    = extract_field(props, '諮詢時間', '諮詢日期', 'date', 'Date', '日期')
    close_d = extract_field(props, '成交日期', 'close_date', 'CloseDate')
    amount  = extract_field(props, '結帳金額（台幣）', '結帳金額', 'amount_twd', 'Amount')

    # Normalize
    name_str   = str(name).strip() if name else ''
    email_str  = normalize_email(email)
    status_str = normalize_status(status)
    date_str   = normalize_date(date)
    close_str  = normalize_date(close_d)
    amount_num = None
    try:
        if amount: amount_num = float(amount)
    except: pass

    # ── Filter: exclude 一人事業 ──────────────────────────────────────────────
    if '一人事業' in name_str:
        skipped_iys += 1
        continue

    # ── Filter: exclude ineffective statuses (unless flag set) ───────────────
    if not args.include_cancelled and status_str in EXCLUDE_STATUS:
        skipped_stat += 1
        continue

    if not email_str:
        # No email = cannot match, skip with warning
        print(f"[警告] 跳過無 email 記錄：{name_str} / 狀態：{status_str}", file=sys.stderr)
        continue

    output.append({
        'email':       email_str,
        'name':        name_str,
        'status':      status_str,
        'date':        date_str,
        'close_date':  close_str,
        'amount_twd':  amount_num,
    })

# ── Save ────────────────────────────────────────────────────────────────────
with open(args.output, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(json.dumps({
    'total_records_in':    len(records),
    'excluded_iys':        skipped_iys,
    'excluded_status':     skipped_stat,
    'valid_records_out':   len(output),
    'output_path':         args.output,
}, ensure_ascii=False, indent=2))
