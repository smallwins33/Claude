"""
BRAND 廣告效益分析腳本
用法：
  python analyze_ads.py --full 廣告報表.csv --7d 7天.csv --4d 4天.csv
                        --leads Systeme名單.csv --consult 諮詢名單.csv
                        --output 輸出.xlsx [--period "2026年3月W13週五"]
                        [--cpl-threshold 4.07]  # 可選，不填則自動計算中位數
"""
import warnings; warnings.filterwarnings("ignore")
import csv, sys, argparse, json
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── CLI ────────────────────────────────────────────────────────────────────────
p = argparse.ArgumentParser()
p.add_argument('--full',    required=True)
p.add_argument('--7d',      required=True, dest='d7')
p.add_argument('--4d',      required=True, dest='d4')
p.add_argument('--leads',     required=True,
               help='Layer 1：本期新名單 CSV（供 leads_by_utm 統計，通常是 systeme_new_leads.csv）')
p.add_argument('--all-leads', default=None, dest='all_leads',
               help='Layer 2：全量名單 CSV（供 email→utm 比對；未指定則沿用 --leads）')
p.add_argument('--consult', default=None)
p.add_argument('--output',  required=True)
p.add_argument('--period',  default='')
p.add_argument('--cpl-threshold', type=float, default=None)
# Notion data: path to JSON file with list of {email, name, status, date}
p.add_argument('--notion',  default=None)
args = p.parse_args()

# ── Helpers ────────────────────────────────────────────────────────────────────
FONT = "Arial"

def load_csv(path, enc="utf-8"):
    try:
        with open(path, encoding=enc) as f:
            return list(csv.DictReader(f))
    except UnicodeDecodeError:
        with open(path, encoding="utf-16") as f:
            return list(csv.DictReader(f))

def sf(v, d=None):
    try: return float(v)
    except: return d

def si(v, d=0):
    try: return int(v)
    except: return d

thin = Side(style="thin", color="CCCCCC")
def tb(): return Border(left=thin, right=thin, top=thin, bottom=thin)
def fl(h): return PatternFill("solid", start_color=h, fgColor=h)
def center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def left(): return Alignment(horizontal="left", vertical="center", wrap_text=True)
def cw(ws, col, w): ws.column_dimensions[get_column_letter(col)].width = w

def wc(ws, row, col, val, bold=False, bg=None, nf=None, align=None, color="000000"):
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(name=FONT, bold=bold, size=10, color=color)
    if bg: c.fill = fl(bg)
    if nf: c.number_format = nf
    c.alignment = align or left()
    c.border = tb()
    return c

def hdr(ws, row, cols, bg="2C3E50"):
    for ci, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=ci, value=h)
        c.font = Font(name=FONT, bold=True, size=10, color="FFFFFF")
        c.fill = fl(bg); c.alignment = center(); c.border = tb()
    ws.row_dimensions[row].height = 36

# ── Load data ──────────────────────────────────────────────────────────────────
ads_full = load_csv(args.full)
ads_7d   = load_csv(args.d7)
ads_4d   = load_csv(args.d4)
# Layer 1：本期新名單（leads_by_utm 統計來源）
systeme  = load_csv(args.leads)
# Layer 2：全量名單（email→utm 比對；未指定時沿用 Layer 1 名單）
systeme_all = load_csv(args.all_leads) if args.all_leads else systeme

consults = []
if args.consult:
    consults = load_csv(args.consult)

notion_data = []
if args.notion:
    with open(args.notion) as f:
        notion_data = json.load(f)

# ── Build lookup structures ────────────────────────────────────────────────────
# Layer 2：email → first Systeme record（first-touch attribution，用全量名單）
email_to_s = {}
for r in sorted(systeme_all, key=lambda x: x.get('Date Registered', '')):
    e = r['电子邮件'].strip().lower()
    if e not in email_to_s:
        email_to_s[e] = r

# Layer 1：utm_content → lead count（用本期新名單）
leads_by_utm = defaultdict(int)
for r in systeme:
    leads_by_utm[r.get('utm_content','').strip()] += 1

# Determine consultation source (Calendly or Notion)
brand_consults = []
if notion_data:
    # Notion takes precedence when available
    for n in notion_data:
        brand_consults.append({
            'email': n.get('email','').strip().lower(),
            'name': n.get('name',''),
            'status': n.get('status',''),   # 成交 / 需跟進 / 未成交
            'date': n.get('date',''),
            'source': 'notion'
        })
elif consults:
    for c in consults:
        if 'B.R.A.N.D' in c.get('Event Type Name',''):
            brand_consults.append({
                'email': c['Invitee Email'].strip().lower(),
                'name': c['Invitee Name'],
                'status': '',
                'date': c.get('Start Date & Time',''),
                'canceled': c.get('Canceled','false'),
                'source': 'calendly'
            })

# utm → consultations  (all statuses)
consult_by_utm = defaultdict(list)
for c in brand_consults:
    s = email_to_s.get(c['email'])
    if s:
        utm = s.get('utm_content','').strip()
        if utm and utm != 'link_in_bio':
            consult_by_utm[utm].append(c)

# utm → conversions (成交 + 需跟進)
def roi_score(c):
    status = c.get('status','')
    if status == '成交': return 1.0
    if status == '需跟進': return 0.5
    if status == '': return 0.5  # unknown (calendly-only) = counts as consult
    return 0.0

def has_roi(utm):
    consults_for_utm = consult_by_utm.get(utm, [])
    if not consults_for_utm:
        return False
    # If using Notion: ROI high = has 成交 or 需跟進
    if notion_data:
        return any(roi_score(c) > 0 for c in consults_for_utm)
    # If Calendly only: ROI high = has any consultation
    return True

# ── CPL threshold ──────────────────────────────────────────────────────────────
ad_7d_map = {a['廣告編號']: a for a in ads_7d}
ad_4d_map = {a['廣告編號']: a for a in ads_4d}

if args.cpl_threshold:
    CPL_THR = args.cpl_threshold
else:
    cpls = sorted([float(a['每次成果成本']) for a in ads_full if a.get('每次成果成本')])
    CPL_THR = cpls[len(cpls)//2] if cpls else 5.0

# ── Build per-ad analysis ──────────────────────────────────────────────────────
def judge(ad_id, cpl_full):
    cpl_low  = (cpl_full is not None) and (cpl_full <= CPL_THR)
    roi_high = has_roi(ad_id)
    a7 = ad_7d_map.get(ad_id, {})
    a4 = ad_4d_map.get(ad_id, {})
    cpl_7d = sf(a7.get('每次成果成本'))
    cpl_4d = sf(a4.get('每次成果成本'))

    if   cpl_low and roi_high:      quad = "⭐ 優等廣告"
    elif not cpl_low and roi_high:  quad = "🔵 第二等"
    elif cpl_low and not roi_high:  quad = "⚠️ 第三等"
    else:                           quad = "🚫 垃圾廣告"

    d4g = (cpl_4d is not None) and (cpl_4d <= CPL_THR)
    d7g = (cpl_7d is not None) and (cpl_7d <= CPL_THR)
    if cpl_4d is None or cpl_7d is None:
        trend = "⚪ 資料不足"; trend_act = "待觀察"
    elif d4g and d7g:             trend = "📈 4天好＋7天好"; trend_act = "加預算"
    elif not d4g and not d7g:     trend = "📉 4天差＋7天差"; trend_act = "立即關閉"
    elif d4g and not d7g:         trend = "📊 4天好＋7天差"; trend_act = "繼續觀察"
    else:                         trend = "🔄 4天差＋7天好"; trend_act = "觀察到下階段"

    n = len(consult_by_utm.get(ad_id, []))
    if n > 0 and trend_act == "立即關閉":
        final = "⚠️ 曾帶來諮詢但CPL惡化，建議暫停優化"
    elif quad == "⭐ 優等廣告" and trend_act == "加預算": final = "✅ 強烈建議加碼"
    elif quad == "⭐ 優等廣告":                          final = "👀 優質素材，持續監控趨勢"
    elif quad == "🚫 垃圾廣告" and trend_act == "立即關閉": final = "❌ 立即關閉"
    elif quad == "🚫 垃圾廣告":                          final = "❌ 建議關閉（無諮詢轉換）"
    elif quad == "⚠️ 第三等":                            final = "⚠️ 謹慎保留，觀察是否帶出諮詢"
    elif quad == "🔵 第二等":                            final = "🔵 保留觀察（CPL高但有轉換潛力）"
    else:                                               final = "⚪ 待觀察"

    return dict(quad=quad, trend=trend, trend_act=trend_act, final=final,
                cpl_7d=cpl_7d, cpl_4d=cpl_4d,
                d4_str="好" if d4g else ("差" if cpl_4d else "—"),
                d7_str="好" if d7g else ("差" if cpl_7d else "—"))

rows = []
for a in ads_full:
    ad_id   = a['廣告編號']
    j       = judge(ad_id, sf(a.get('每次成果成本')))
    a7      = ad_7d_map.get(ad_id, {})
    a4      = ad_4d_map.get(ad_id, {})
    c_list  = consult_by_utm.get(ad_id, [])
    # Create dict from judge() but exclude cpl_7d and cpl_4d since we handle them below
    j_clean = {k: v for k, v in j.items() if k not in ['cpl_7d', 'cpl_4d']}
    rows.append(dict(
        ad_id=ad_id, ad_name=a['廣告名稱'], ad_group=a['廣告組合名稱'],
        ad_status=a.get('廣告投遞',''),
        cpl_full=sf(a.get('每次成果成本')), spend_full=sf(a.get('花費金額 (USD)')),
        leads_full=si(a.get('成果')),
        reach=si(a.get('觸及人數')), impressions=si(a.get('曝光次數')),
        cpl_7d=j['cpl_7d'], spend_7d=sf(a7.get('花費金額 (USD)')), leads_7d=si(a7.get('成果')),
        cpl_4d=j['cpl_4d'], spend_4d=sf(a4.get('花費金額 (USD)')), leads_4d=si(a4.get('成果')),
        s_leads=leads_by_utm.get(ad_id, 0),
        n_consult=len(c_list),
        n_converted=sum(1 for c in c_list if c.get('status') == '成交'),
        n_followup=sum(1 for c in c_list if c.get('status') == '需跟進'),
        consult_names="、".join(c['name'] for c in c_list) or "—",
        **j_clean
    ))

# ── Color maps ────────────────────────────────────────────────────────────────
QUAD_C = {
    "⭐ 優等廣告": ("D4EDDA","1E7E34"),
    "🔵 第二等":   ("D1ECF1","0C5460"),
    "⚠️ 第三等":   ("FFF3CD","856404"),
    "🚫 垃圾廣告": ("F8D7DA","721C24"),
}
TREND_C = {
    "📈 4天好＋7天好": ("D4EDDA","1E7E34"),
    "📊 4天好＋7天差": ("FFF3CD","856404"),
    "🔄 4天差＋7天好": ("D1ECF1","0C5460"),
    "📉 4天差＋7天差": ("F8D7DA","721C24"),
    "⚪ 資料不足":     ("F8F9FA","6C757D"),
}
FINAL_C = {
    "✅ 強烈建議加碼":              ("D4EDDA","155724"),
    "👀 優質素材，持續監控趨勢":    ("D1ECF1","0C5460"),
    "⚠️ 曾帶來諮詢但CPL惡化，建議暫停優化": ("FFF3CD","856404"),
    "⚠️ 謹慎保留，觀察是否帶出諮詢": ("FFF3CD","856404"),
    "🔵 保留觀察（CPL高但有轉換潛力）": ("D1ECF1","0C5460"),
    "❌ 立即關閉":                  ("F8D7DA","721C24"),
    "❌ 建議關閉（無諮詢轉換）":    ("F8D7DA","721C24"),
    "⚪ 待觀察":                    ("F8F9FA","6C757D"),
}

# ── Build Excel ────────────────────────────────────────────────────────────────
wb = Workbook()

# ── Sheet 1: 廣告素材效益 ───────────────────────────────────────────────────────
ws1 = wb.active
ws1.title = "📢 廣告素材效益"
ws1.sheet_view.showGridLines = False

title = f"BRAND 廣告素材效益分析｜{args.period}｜CPL 閾值 ${CPL_THR:.2f}"
ws1.merge_cells(f"A1:R1")
ws1["A1"].value = title
ws1["A1"].font = Font(name=FONT, bold=True, size=13, color="FFFFFF")
ws1["A1"].fill = fl("1A2E44"); ws1["A1"].alignment = center()
ws1.row_dimensions[1].height = 34

use_notion = bool(notion_data)
hdrs1 = ["廣告名稱","廣告編號","廣告組合","狀態",
         "全月\nMeta成果","全月\nCPL","Systeme\n名單數","Systeme\nCPL",
         "全月花費\n(USD)","觸及\n人數","曝光\n次數",
         "諮詢\n數量","諮詢者姓名",
         "四象限","趨勢\n(4/7天)","趨勢\n建議","最終建議"]
if use_notion:
    hdrs1 = hdrs1[:12] + ["成交\n數量","需跟進\n數量"] + hdrs1[12:]

hdr(ws1, 2, hdrs1)
total_row = len(rows) + 3

for ri, r in enumerate(rows, 3):
    ws1.row_dimensions[ri].height = 22
    is_alt = ri % 2 == 0
    base_bg = "EBF3FB" if is_alt else "FFFFFF"
    qbg, qfg = QUAD_C.get(r['quad'], (base_bg,"000000"))
    tbg, tfg = TREND_C.get(r['trend'], (base_bg,"000000"))
    fbg, ffg = FINAL_C.get(r['final'], (base_bg,"000000"))

    sys_cpl = r['spend_full']/r['s_leads'] if r['s_leads'] else None
    base_vals = [
        r['ad_name'], r['ad_id'], r['ad_group'], r['ad_status'],
        r['leads_full'], r['cpl_full'],
        r['s_leads'] if r['s_leads'] else "—", sys_cpl,
        r['spend_full'], r['reach'], r['impressions'],
        r['n_consult'], r['consult_names'],
    ]
    base_nfs = [None,None,None,None,"#,##0","$#,##0.00","#,##0","$#,##0.00",
                "$#,##0.00","#,##0","#,##0","#,##0",None]
    if use_notion:
        base_vals = base_vals[:12] + [r['n_converted'], r['n_followup']] + base_vals[12:]
        base_nfs  = base_nfs[:12]  + ["#,##0","#,##0"] + base_nfs[12:]

    decision_vals = [r['quad'], r['trend'], r['trend_act'], r['final']]
    decision_bgs  = [qbg, tbg, tbg, fbg]
    decision_fgs  = [qfg, tfg, tfg, ffg]

    all_vals = base_vals + decision_vals
    all_nfs  = base_nfs  + [None,None,None,None]
    n_base   = len(base_vals)

    for ci, (v, nf) in enumerate(zip(all_vals, all_nfs), 1):
        di = ci - n_base - 1
        if di >= 0:
            wc(ws1, ri, ci, v, bold=True, bg=decision_bgs[di], nf=nf,
               color=decision_fgs[di], align=center())
        else:
            wc(ws1, ri, ci, v, bg=base_bg, nf=nf)

# Totals row
n_cols = len(hdrs1)
ws1.row_dimensions[total_row].height = 24
wc(ws1, total_row, 1, "合計", bold=True, bg="2C3E50", color="FFFFFF")
for ci in range(2, n_cols+1):
    wc(ws1, total_row, ci, "", bg="2C3E50")
# Sum meta leads (col 5), spend (col 9)
meta_col = 5; spend_col = 9; consult_col = 12
ws1.cell(total_row, meta_col).value  = f"=SUM({get_column_letter(meta_col)}3:{get_column_letter(meta_col)}{total_row-1})"
ws1.cell(total_row, spend_col).value = f"=SUM({get_column_letter(spend_col)}3:{get_column_letter(spend_col)}{total_row-1})"
ws1.cell(total_row, consult_col).value = f"=SUM({get_column_letter(consult_col)}3:{get_column_letter(consult_col)}{total_row-1})"
for ci in [meta_col, spend_col, consult_col]:
    ws1.cell(total_row, ci).font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
    ws1.cell(total_row, ci).number_format = "$#,##0.00" if ci == spend_col else "#,##0"

col_widths1 = [28,22,28,8,10,12,10,12,12,10,10,8,20,16,20,14,28]
if use_notion: col_widths1 = col_widths1[:12] + [8,8] + col_widths1[12:]
for ci, w in enumerate(col_widths1, 1): cw(ws1, ci, w)

# ── Sheet 2: 諮詢比對明細 ───────────────────────────────────────────────────────
ws2 = wb.create_sheet("📋 諮詢比對明細")
ws2.sheet_view.showGridLines = False
ws2.merge_cells("A1:H1")
ws2["A1"].value = f"BRAND 諮詢來源比對明細｜{args.period}"
ws2["A1"].font = Font(name=FONT, bold=True, size=13, color="FFFFFF")
ws2["A1"].fill = fl("1A2E44"); ws2["A1"].alignment = center()
ws2.row_dimensions[1].height = 34

h2 = ["諮詢時間","姓名","Email","在Systeme？","utm_content","廣告名稱","廣告花費(USD)"]
if use_notion:
    h2 += ["成交狀態"]
h2 += ["來源分類"]
hdr(ws2, 2, h2)

ad_name_map = {a['廣告編號']: a['廣告名稱'] for a in ads_full}
ad_spend_map = {a['廣告編號']: sf(a.get('花費金額 (USD)')) for a in ads_full}

# Build full consult detail list
cat_colors = {
    "✅ Meta廣告（可追蹤）": ("D4EDDA","1E7E34"),
    "🔗 Link in Bio":        ("FFF3CD","856404"),
    "❓ 無UTM（其他）":       ("FEF9E7","5A4E00"),
    "❌ 不在Systeme名單":     ("F8D7DA","721C24"),
}
detail_rows = []
for c in brand_consults:
    e = c['email']
    s = email_to_s.get(e)
    utm = s.get('utm_content','').strip() if s else None
    ad_n = ad_name_map.get(utm,'') if utm else ''
    ad_sp = ad_spend_map.get(utm) if utm else None
    if not s:
        cat = "❌ 不在Systeme名單"
    elif utm and utm != 'link_in_bio' and ad_n:
        cat = "✅ Meta廣告（可追蹤）"
    elif utm == 'link_in_bio':
        cat = "🔗 Link in Bio"
    else:
        cat = "❓ 無UTM（其他）"
    detail_rows.append((c, utm, ad_n, ad_sp, cat))

sort_order = {"✅ Meta廣告（可追蹤）":0,"🔗 Link in Bio":1,"❓ 無UTM（其他）":2,"❌ 不在Systeme名單":3}
detail_rows.sort(key=lambda x: sort_order.get(x[4], 9))

for ri, (c, utm, ad_n, ad_sp, cat) in enumerate(detail_rows, 3):
    ws2.row_dimensions[ri].height = 22
    bg, fg = cat_colors.get(cat, ("FFFFFF","000000"))
    vals = [c.get('date',''), c['name'], c['email'],
            "是" if email_to_s.get(c['email']) else "否",
            utm or "（無）", ad_n or "（無）", ad_sp]
    if use_notion:
        vals.append(c.get('status','—'))
    vals.append(cat)
    nfs = [None,None,None,None,None,None,"$#,##0.00"] + ([None] if use_notion else []) + [None]
    for ci, (v, nf) in enumerate(zip(vals, nfs), 1):
        is_last = ci == len(vals)
        wc(ws2, ri, ci, v, bold=is_last, bg=bg, nf=nf,
           color=fg if is_last else "000000")

for ci, w in enumerate([20,14,30,8,22,26,12]+([12] if use_notion else [])+[24], 1):
    cw(ws2, ci, w)

# ── Sheet 3: Systeme 名單統計 ──────────────────────────────────────────────────
ws3 = wb.create_sheet("📥 Systeme名單統計")
ws3.sheet_view.showGridLines = False
ws3.merge_cells("A1:G1")
ws3["A1"].value = f"Systeme 名單 UTM 統計｜{args.period}"
ws3["A1"].font = Font(name=FONT, bold=True, size=13, color="FFFFFF")
ws3["A1"].fill = fl("1A2E44"); ws3["A1"].alignment = center()
ws3.row_dimensions[1].height = 34

hdr(ws3, 2, ["utm_content","廣告名稱","廣告花費(USD)","Meta成果數","Systeme名單數","諮詢數","諮詢轉換率"])

all_utms = sorted(leads_by_utm.keys(), key=lambda x: -leads_by_utm[x])
for ri, utm in enumerate(all_utms, 3):
    ws3.row_dimensions[ri].height = 22
    ad = next((a for a in ads_full if a['廣告編號'] == utm), None)
    count = leads_by_utm[utm]
    n_c = len(consult_by_utm.get(utm, []))
    bg = "D4EDDA" if n_c > 0 else ("EBF3FB" if ri%2==0 else "FFFFFF")
    if utm == 'link_in_bio':
        an, asp, ml = "（Link in Bio）", None, None
    elif ad:
        an, asp, ml = ad['廣告名稱'], sf(ad.get('花費金額 (USD)')), si(ad.get('成果'))
    else:
        an, asp, ml = "（廣告報表中無對應）", None, None
    rate = n_c/count if count else 0
    vals = [utm or "（空白）", an, asp, ml, count, n_c, rate]
    nfs  = [None,None,"$#,##0.00","#,##0","#,##0","#,##0","0.0%"]
    for ci, (v, nf) in enumerate(zip(vals, nfs), 1):
        wc(ws3, ri, ci, v, bold=(ci==6 and n_c>0), bg=bg, nf=nf,
           color="1E7E34" if (ci==6 and n_c>0) else "000000")

for ci, w in enumerate([26,28,12,12,12,8,12], 1): cw(ws3, ci, w)

# ── Save ───────────────────────────────────────────────────────────────────────
wb.save(args.output)
print(json.dumps({
    "output": args.output, "cpl_threshold": CPL_THR,
    "total_ads": len(rows),
    "total_brand_consults": len(brand_consults),
    "summary": {
        "強烈加碼": sum(1 for r in rows if r['final']=="✅ 強烈建議加碼"),
        "優質監控": sum(1 for r in rows if r['final']=="👀 優質素材，持續監控趨勢"),
        "立即關閉": sum(1 for r in rows if r['final']=="❌ 立即關閉"),
        "建議關閉": sum(1 for r in rows if "建議關閉" in r['final']),
        "特殊觀察": sum(1 for r in rows if "曾帶來諮詢" in r['final']),
        "謹慎保留": sum(1 for r in rows if "謹慎保留" in r['final']),
    }
}, ensure_ascii=False, indent=2))
