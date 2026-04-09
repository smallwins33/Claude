"""
fetch_systeme.py
透過 Systeme MCP server 抓取聯絡人名單（含 utm_content）

Layer 1（廣告成效，只抓本期）：
  python scripts/fetch_systeme.py --mode new --since 2026-03-31
  → 使用 registeredAfter API 參數，伺服器端篩選，快速

Layer 2（轉換追蹤，本期前的歷史名單）：
  python scripts/fetch_systeme.py --mode full [--tags 1146739]
  → 全量或依 tag 篩選
"""

import warnings
warnings.filterwarnings("ignore")

import csv
import json
import os
import sys
from datetime import date
from typing import Optional
from urllib.parse import urlparse, parse_qs
import requests
from dotenv import load_dotenv

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), "..", ".env"))

MCP_KEY = os.getenv("SYSTEME_MCP_KEY")
MCP_URL = "https://mcp.systeme.io/mcp"

if not MCP_KEY:
    print("ERROR: SYSTEME_MCP_KEY 未設定", file=sys.stderr)
    sys.exit(1)


# ── MCP 連線 ──────────────────────────────────────────────────────────────────

def init_session() -> str:
    resp = requests.post(
        f"{MCP_URL}?mcpKey={MCP_KEY}",
        headers={"Content-Type": "application/json", "Accept": "application/json"},
        json={
            "jsonrpc": "2.0", "method": "initialize",
            "params": {"protocolVersion": "2024-11-05", "capabilities": {},
                       "clientInfo": {"name": "ads-agent", "version": "2.1"}},
            "id": 1,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.headers["Mcp-Session-Id"]


def call_tool(session_id: str, tool: str, arguments: dict, req_id: int) -> dict:
    import time
    for attempt in range(5):
        resp = requests.post(
            f"{MCP_URL}?mcpKey={MCP_KEY}",
            headers={"Content-Type": "application/json", "Accept": "application/json",
                     "Mcp-Session-Id": session_id},
            json={"jsonrpc": "2.0", "method": "tools/call",
                  "params": {"name": tool, "arguments": arguments}, "id": req_id},
            timeout=30,
        )
        if resp.status_code == 429:
            wait = 2 ** attempt
            print(f"\n  [429] Rate limit，等 {wait}s 後重試（第 {attempt+1} 次）...", end="")
            time.sleep(wait)
            continue
        resp.raise_for_status()
        result = resp.json()
        if "error" in result:
            raise RuntimeError(f"MCP error: {result['error']}")
        content = result["result"]["content"]
        return json.loads(content[0]["text"] if content else "{}")
    raise RuntimeError("call_tool 超過最大重試次數（429）")


# ── 工具函式 ──────────────────────────────────────────────────────────────────

def extract_utm_content(source_url: Optional[str]) -> str:
    if not source_url:
        return ""
    try:
        qs = parse_qs(urlparse(source_url).query)
        values = qs.get("utm_content", [])
        return values[0] if values else ""
    except Exception:
        return ""


# ── 核心抓取 ──────────────────────────────────────────────────────────────────

def fetch_contacts(session_id: str, mode: str = "full",
                   since_date: Optional[date] = None,
                   tags: Optional[str] = None) -> list:
    """
    mode='new'  : registeredAfter=since_date（Layer 1，伺服器端篩選）
    mode='full' : 全量或依 tags 篩選（Layer 2）

    防 loop：seen_ids 追蹤，遇重複 ID 立即停止。
    """
    if mode == "new" and not since_date:
        raise ValueError("mode=new 需要同時指定 --since DATE")

    label = f"本期新名單（registeredAfter={since_date}）" if mode == "new" else "全量"
    if tags:
        label += f"（tags={tags}）"
    print(f"開始抓取 Systeme 聯絡人 [{label}]...")

    contacts = []
    seen_ids: set = set()
    starting_after = None
    req_id = 10

    while True:
        # mode=new 用 order:desc（最新→最舊），遇舊日期手動停止
        # mode=full 用 order:asc
        order = "desc" if mode == "new" else "asc"
        params: dict = {"limit": 100, "order": order}
        if starting_after:
            params["startingAfter"] = starting_after
        if tags:
            params["tags"] = tags

        data = call_tool(session_id, "get_contacts", {"data": params}, req_id)
        req_id += 1

        items = data.get("items", [])
        if not items:
            break

        # ── 防 loop：檢查是否拿到重複 ID ──────────────────────────────────────
        first_id = items[0]["id"]
        if first_id in seen_ids:
            print(f"\n  [防 loop] 偵測到重複 ID {first_id}，停止分頁。")
            break

        for item in items:
            seen_ids.add(item["id"])

        if mode == "new" and since_date:
            # 手動過濾：丟掉早於 since_date 的記錄，遇到就停止
            new_items = []
            stop = False
            for item in items:
                reg = item.get("registeredAt") or item.get("createdAt", "")
                if reg:
                    try:
                        from datetime import datetime
                        item_date = datetime.fromisoformat(
                            reg.replace("Z", "+00:00")).date()
                        if item_date < since_date:
                            stop = True
                            break
                    except Exception:
                        pass
                new_items.append(item)
            contacts.extend(new_items)
            print(f"  已抓取 {len(contacts)} 筆（本期）...", end="\r")
            if stop:
                print(f"\n  遇到 {since_date} 前的記錄，停止。")
                break
        else:
            contacts.extend(items)
            print(f"  已抓取 {len(contacts)} 筆...", end="\r")

        if len(items) < 100:
            break

        starting_after = items[-1]["id"]

    print(f"\n共抓取 {len(contacts)} 筆聯絡人（唯一 ID）")
    return contacts


def extract_leads(contacts: list) -> list:
    leads = []
    for c in contacts:
        email = c.get("email", "").strip().lower()
        if not email:
            continue
        leads.append({
            "email": email,
            "utm_content": extract_utm_content(c.get("sourceURL")),
            "registered_at": c.get("registeredAt") or c.get("createdAt", ""),
            "systeme_id": c.get("id"),
        })
    return leads


def save_outputs(leads: list, json_path: str):
    csv_path = json_path.replace(".json", ".csv")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(leads, f, ensure_ascii=False, indent=2)
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=["电子邮件", "utm_content", "Date Registered", "systeme_id"])
        writer.writeheader()
        for lead in leads:
            writer.writerow({
                "电子邮件": lead["email"],
                "utm_content": lead["utm_content"],
                "Date Registered": lead["registered_at"],
                "systeme_id": lead["systeme_id"],
            })
    print(f"已儲存 {len(leads)} 筆 → {json_path} 及 {csv_path}")


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--mode", choices=["new", "full"], default="full")
    p.add_argument("--since", default=None, help="mode=new 必填：YYYY-MM-DD")
    p.add_argument("--tags", default=None,
                   help="tag ID（逗號分隔），e.g. 1146739，不填則不篩選")
    p.add_argument("--out", default=None)
    args = p.parse_args()

    since_date = None
    if args.mode == "new":
        if not args.since:
            print("ERROR: mode=new 必須指定 --since YYYY-MM-DD", file=sys.stderr)
            sys.exit(1)
        since_date = date.fromisoformat(args.since)

    out_dir = os.path.join(os.path.dirname(__file__), "..", "output")
    os.makedirs(out_dir, exist_ok=True)
    default_out = (os.path.join(out_dir, "systeme_new_leads.json") if args.mode == "new"
                   else os.path.join(out_dir, "systeme_leads.json"))
    output_path = args.out or default_out

    session_id = init_session()
    print(f"MCP Session 建立成功（mode={args.mode}）")

    contacts = fetch_contacts(session_id, mode=args.mode,
                              since_date=since_date, tags=args.tags)
    leads = extract_leads(contacts)
    save_outputs(leads, output_path)

    utm_counts: dict = {}
    for lead in leads:
        utm = lead["utm_content"] or "(無 utm)"
        utm_counts[utm] = utm_counts.get(utm, 0) + 1
    print(f"\nutm_content 分布（前 20）：")
    for utm, count in sorted(utm_counts.items(), key=lambda x: -x[1])[:20]:
        print(f"  {utm}: {count} 筆")

    print(json.dumps({
        "mode": args.mode,
        "since": str(since_date) if since_date else None,
        "tags": args.tags,
        "total_leads": len(leads),
        "output_json": output_path,
        "output_csv": output_path.replace(".json", ".csv"),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
