"""
ChangeGear 派單原則歸納工具（每週反思）
========================================

從過去 N 天的人工修正（feedback + corrected=1 工單）中，
呼叫 Claude 歸納出高層派單原則，寫入 learned_principles.md。
bot 下次派單時會把這些原則注入 Claude system prompt。

這是「深度學習」的核心：不只記得個別案例（teach.py），
還會週期性地從多個案例中提煉出可推廣的原則。

執行方式：
    python reflect.py              # 分析過去 7 天，產出原則
    python reflect.py --days 30    # 分析過去 30 天
    python reflect.py --dry-run    # 印到 console，不寫檔
    python reflect.py --show       # 顯示目前的 learned_principles.md
"""

import sqlite3
import sys
import json
import re
import os
from datetime import datetime, timedelta

import openpyxl
import anthropic

DB_PATH = "changegear_history.db"
PRINCIPLES_FILE = "learned_principles.md"
HISTORY_DIR = "learned_principles_history"


# ── Config / Material ──────────────────────────────────────────
def load_config() -> dict:
    try:
        wb = openpyxl.load_workbook("ChangeGear_AutoAssign_Rules.xlsx")
        ws = wb["程式設定"]
        cfg = {str(r[0]).strip(): str(r[1]).strip()
               for r in ws.iter_rows(min_row=3, values_only=True)
               if r[0] and r[1]}
        return {
            "api_key": cfg.get("Claude API Key", ""),
            "model":   cfg.get("Claude 模型",   "claude-sonnet-4-5"),
        }
    except Exception as e:
        print(f"⛔ 讀取 Excel 失敗: {e}")
        sys.exit(1)


def gather_material(days: int = 7) -> dict:
    cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(DB_PATH)
    try:
        # feedback 是顯式教學
        feedback = conn.execute("""
            SELECT ticket_id, summary, requester, req_item,
                   wrong_owner, wrong_assigned_to, correct_owner,
                   correct_assigned_to, reason, created_at
            FROM feedback
            WHERE created_at >= ?
            ORDER BY created_at DESC
        """, (cutoff,)).fetchall()

        # corrected=1 工單是隱式修正（bot 派錯被人工改過）
        corrections = conn.execute("""
            SELECT ticket_id, summary, requester, owner, assigned_to,
                   inc_parent, inc_child, inc_item, req_item, updated_at
            FROM assignments
            WHERE corrected = 1 AND updated_at >= ?
            ORDER BY updated_at DESC
            LIMIT 100
        """, (cutoff,)).fetchall()
    finally:
        conn.close()

    return {
        "feedback":     feedback,
        "corrections":  corrections,
        "cutoff_date":  cutoff,
        "days":         days,
    }


# ── Prompt ──────────────────────────────────────────────────────
def build_prompt(material: dict) -> str:
    feedback_text = ""
    if material["feedback"]:
        feedback_text = "## 教學紀錄（操作員主動標記）\n"
        for i, r in enumerate(material["feedback"], 1):
            tid, summary, req, ri, wo, wa, co, ca, reason, ts = r
            feedback_text += (
                f"[{i}] {tid} @ {ts}\n"
                f"  摘要:     {summary}\n"
                f"  寄件者:   {req}\n"
                f"  Req Item: {ri}\n"
                f"  ✗ 錯誤:   Owner={wo} / Assigned={wa}\n"
                f"  ✓ 正確:   Owner={co} / Assigned={ca}\n"
                f"  原因:     {reason}\n\n"
            )

    correction_text = ""
    if material["corrections"]:
        correction_text = "## 隱式修正（bot 派錯後被人工改）\n"
        for i, r in enumerate(material["corrections"], 1):
            tid, summary, req, owner, assigned, p, c, it, ri, ts = r
            correction_text += (
                f"[{i}] {tid} @ {ts}\n"
                f"  摘要:        {summary}\n"
                f"  寄件者:      {req}\n"
                f"  正確派給:    Owner={owner} / Assigned={assigned}\n"
                f"  Inc Type:    {p} > {c} > {it}\n"
                f"  Req Item:    {ri}\n\n"
            )

    return f"""你是 IT 服務台派單顧問。請分析過去 {material['days']} 天的派單修正案例，
歸納出一份「高層原則文件」給自動派單 bot 未來判斷類似工單時使用。

{feedback_text}
{correction_text}

## 任務
辨識重複出現的派單模式，歸納出 3–10 條原則。每條原則必須具備：
  - title:            一句話原則名稱
  - keywords:         識別此情境的關鍵字或 req_item 樣式（陣列；至少 1 個）
  - inc_type:         適用的 Incident Type 樣式（可留空）
  - correct_assignee: 應派的個人（AD 帳號或顯示名稱）
  - correct_owner:    應派的團隊（如 Help Desk / Apps Team）
  - rationale:        為什麼這樣派（一句話）
  - example_tickets:  參考案例的 ticket_id 陣列（最多 3 個）

## 重要
- 只歸納「重複出現至少 2 次」或「reason 強烈指出特定派單規則」的原則
- 不要把單一案例硬升級成原則
- 若資料不足以歸納，回傳 {{"principles": []}}

僅回傳 JSON，格式：
{{
  "principles": [
    {{"title": "...", "keywords": [...], "inc_type": "...",
      "correct_assignee": "...", "correct_owner": "...",
      "rationale": "...", "example_tickets": [...]}}
  ]
}}
不要回傳 markdown code block，直接 JSON。"""


# ── Claude call ─────────────────────────────────────────────────
def call_claude(prompt: str, api_key: str, model: str) -> list:
    client = anthropic.Anthropic(api_key=api_key)
    resp = client.messages.create(
        model=model,
        max_tokens=4096,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = resp.content[0].text.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    try:
        data = json.loads(raw)
        return data.get("principles", [])
    except json.JSONDecodeError as e:
        print(f"⛔ Claude 回傳 JSON 解析失敗: {e}")
        print(f"原始回傳: {raw[:500]}")
        return []


# ── Markdown output ─────────────────────────────────────────────
def format_markdown(principles: list, material: dict) -> str:
    if not principles:
        return (
            f"# 派單原則（自動歸納）\n\n"
            f"*過去 {material['days']} 天無足夠案例可歸納（"
            f"feedback={len(material['feedback'])} 筆 / "
            f"corrections={len(material['corrections'])} 筆）*\n"
        )

    md = f"""# 派單原則（Claude 自動歸納）

> 由過去 {material['days']} 天的 {len(material['feedback'])} 筆教學紀錄
> 與 {len(material['corrections'])} 筆隱式修正歸納而來
> 產出時間：{datetime.now().strftime('%Y-%m-%d %H:%M')}

bot 在派單時會將此檔案內容注入 Claude system prompt，
作為高層判斷依據（優先級高於三信號決策）。

"""
    for i, p in enumerate(principles, 1):
        kws = ", ".join(p.get("keywords", []))
        md += f"## 原則 {i}：{p.get('title', '(無標題)')}\n\n"
        md += f"- **關鍵字 / 樣式**: {kws or '(無)'}\n"
        if p.get("inc_type"):
            md += f"- **Incident Type**: {p['inc_type']}\n"
        md += f"- **應派 Owner**: {p.get('correct_owner', '(未指定)')}\n"
        md += f"- **應派 Assigned**: {p.get('correct_assignee', '(未指定)')}\n"
        md += f"- **判斷依據**: {p.get('rationale', '(無)')}\n"
        ex = p.get("example_tickets", [])
        if ex:
            md += f"- **參考案例**: {', '.join(ex)}\n"
        md += "\n"
    return md


# ── Main ────────────────────────────────────────────────────────
def main():
    args = sys.argv[1:]
    days = 7
    dry_run = False
    show_only = False
    while args:
        a = args.pop(0)
        if a == "--days" and args:
            days = int(args.pop(0))
        elif a == "--dry-run":
            dry_run = True
        elif a == "--show":
            show_only = True
        elif a in ("--help", "-h"):
            print(__doc__)
            return

    if show_only:
        if os.path.exists(PRINCIPLES_FILE):
            with open(PRINCIPLES_FILE, encoding="utf-8") as f:
                print(f.read())
        else:
            print(f"⛔ {PRINCIPLES_FILE} 不存在，請先跑一次 reflect.py")
        return

    print("=" * 60)
    print(f"ChangeGear 派單原則歸納（過去 {days} 天）")
    print("=" * 60)

    cfg = load_config()
    if not cfg["api_key"]:
        print("⛔ Excel 內沒設定 Claude API Key")
        sys.exit(1)

    print(f"\n[1/3] 蒐集學習素材...")
    material = gather_material(days)
    fb_n = len(material["feedback"])
    co_n = len(material["corrections"])
    print(f"  feedback 教學: {fb_n} 筆")
    print(f"  隱式修正:     {co_n} 筆")

    if fb_n == 0 and co_n == 0:
        print(f"\n⚠ 過去 {days} 天內無任何學習素材，本次不更新原則檔")
        return

    print(f"\n[2/3] 呼叫 Claude ({cfg['model']}) 歸納原則...")
    prompt = build_prompt(material)
    principles = call_claude(prompt, cfg["api_key"], cfg["model"])
    print(f"  歸納出 {len(principles)} 條原則")

    md = format_markdown(principles, material)

    if dry_run:
        print("\n[3/3] --dry-run 模式，輸出到 console：\n")
        print(md)
        return

    print(f"\n[3/3] 寫入 {PRINCIPLES_FILE}")
    if os.path.exists(PRINCIPLES_FILE):
        os.makedirs(HISTORY_DIR, exist_ok=True)
        week = datetime.now().strftime("%Y-W%V")
        backup = os.path.join(HISTORY_DIR, f"learned_principles_{week}.md")
        # 若同一週已備份過，加時間戳避免覆蓋
        if os.path.exists(backup):
            backup = backup.replace(".md", f"_{datetime.now().strftime('%H%M')}.md")
        os.rename(PRINCIPLES_FILE, backup)
        print(f"  舊版備份至: {backup}")

    with open(PRINCIPLES_FILE, "w", encoding="utf-8") as f:
        f.write(md)

    print(f"\n✅ 完成。bot 下次派單時會自動讀取 {PRINCIPLES_FILE}")
    print(f"   查看：python reflect.py --show")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n已中止")
        sys.exit(0)
