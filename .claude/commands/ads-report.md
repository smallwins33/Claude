# 廣告成效報告

領先時代 B.R.A.N.D 課程廣告報表。每週兩次：**週一大盤點**、**週五快檢**。

直接執行，不需要問任何問題。

---

## 週一大盤點（每週一執行）

**報告期間**：上週一 ～ 上週日

### 執行步驟

**Step 1｜拉 Meta 廣告數據**
```
python3 scripts/fetch_meta.py --since LAST_MON --until LAST_SUN \
  --out-full output/meta_full.csv \
  --out-7d output/meta_7d.csv \
  --out-4d output/meta_4d.csv
```

**Step 2｜拉 Systeme 名單**
```
python3 scripts/fetch_systeme.py --mode new --since LAST_MON --out output/systeme_lta_MMDD.csv
```
⚠️ registeredAfter / tags filter 無效，一律 order:desc + seen_ids 防重複，遇到 since 前的記錄立即停止。

**Step 3｜拉 Notion 諮詢記錄**
- 用 `notion-query-data-sources`，資料庫：`collection://346b8e54-61be-49a4-8379-2a1f6a33075b`
- 只取 `諮詢時間 >= LAST_MON`，不抓全量
- 欄位：Email、客戶名稱、狀態、date:諮詢時間:start
- **永遠不呼叫 notion-fetch**
- 存成 `output/notion_consult_MMDD.json`，格式：`[{email, name, status, date}]`

**Step 4｜產出 Excel 報表**
```
python3 scripts/analyze_ads.py \
  --full output/meta_full.csv \
  --7d output/meta_7d.csv \
  --4d output/meta_4d.csv \
  --leads output/systeme_lta_MMDD.csv \
  --notion output/notion_consult_MMDD.json \
  --output output/ads_report_lta_MMDD.xlsx \
  --period "YYYY年M月WXX上週"
```

**Step 5｜輸出文字摘要**（格式見下方）

---

## 週五快檢（每週五執行）

**報告期間**：本週一 ～ 昨天（週四）
**目的**：只看 CPL 有沒有超標，不做 ROI 判定（太早，數據未成熟）

```
python3 scripts/fetch_meta.py --since THIS_MON --until YESTERDAY \
  --out-full output/meta_midweek.csv \
  --out-7d output/meta_midweek_7d.csv \
  --out-4d output/meta_midweek_4d.csv
```

快檢只輸出文字，不產 Excel。格式：
```
【週五快檢】YYYY-MM-DD（週一）～ YYYY-MM-DD（週四）

⚠️ CPL 警示（$5–8）：
- 廣告名稱｜CPL $X.XX｜趨勢：...

❌ CPL 超標（≥$8）：
- 廣告名稱｜CPL $X.XX｜建議：立即暫停

✅ 其餘廣告 CPL 正常（≤$5），無異常。
```

---

## 判定機制

### 廣告歸因
- **廣告編號** = Systeme 裡的 `utm_content` 欄位，即 Meta 廣告 ID（數字串）
- 不叫 UTM，叫廣告編號

### 核心指標

**CPQL（Cost Per Qualified Lead）= 花費 ÷ 總諮詢數（含未出席）**
- 反映廣告帶來諮詢的整體成本

**CP合格QL = 花費 ÷ 有出席諮詢數（排除未出席）**
- 未出席 = 動機不足，視為非受眾
- 排除後更準確反映廣告帶到有意願受眾的成本
- 這是更重要的品質指標

- CPL 為輔助指標（控制獲客量與成本效率）

### Stage 1｜廣告責任（四象限）

**CPL 三段標準**
| 範圍 | 判定 |
|---|---|
| ≤ $5 USD | ✅ 好 |
| $5–8 USD | ⚠️ 警示 |
| ≥ $8 USD | ❌ 不行 |

**ROI 確認**
- `has_roi` = 該廣告編號有任何一筆諮詢預約，不論成交與否
- 「有進諮詢」即算 ROI 確認（已被漏斗預篩，品質可信）

**四象限**
| | CPL 好（≤$5） | CPL 警示（$5–8） | CPL 不行（≥$8） |
|---|---|---|---|
| 有諮詢 | ⭐ 優等廣告 | 🔵 第二等 | 🟠 高CPL有轉換 |
| 無諮詢 | ⚠️ 第三等 | 🟡 警示觀察 | 🚫 垃圾廣告 |

**停廣告門檻**（需同時符合）
1. CPL ≥ $8
2. 近 4 天 + 近 7 天趨勢都差

### 出席率監控（唯一品質信號）

- 未出席率 ≥ 50% 時觸發警示
- 代表廣告文案與課程期待有落差，建議檢視素材
- 不直接砍廣告，做素材檢討

**注意：** 不做複雜的 Stage 2 名單品質判斷——漏斗本身（15分鐘等待 + 問卷 + 預算確認）已預先篩選，進到諮詢的人品質由漏斗保證。

---

## 週一大盤點文字摘要格式

```
【廣告期間】YYYY-MM-DD ～ YYYY-MM-DD

【名單】
- 總名單數：N 筆
- 無廣告編號名單：N 筆（X%）
- CPL：$X.XX USD（若無預算資料標注「待補」）

【諮詢】
- 期間諮詢數：N 筆
- 可歸因廣告的諮詢：N 筆

【各廣告判定】
- ⭐ 優等：X 支
- 🔵 第二等：X 支
- ⚠️ 第三等：X 支
- 🟡 警示：X 支
- 🚫 垃圾：X 支

【Stage 2 警示】（有觸發才列，否則省略）
- 廣告名稱｜諮詢 N 筆｜有效率 X%｜建議檢討素材角度

【優化建議】
（廣告策略師角度，根據實際數字）
```

---

## 品牌資訊

- **品牌**：領先時代 B.R.A.N.D 課程
- **Notion 諮詢 DB**：`collection://346b8e54-61be-49a4-8379-2a1f6a33075b`
- **輸出路徑**：`/Users/mi/Developer/Claude/output/`
- **廣告平台**：Facebook / Instagram
