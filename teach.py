"""
ChangeGear 派單修正回饋工具
============================

互動式工具：告訴 bot 哪張工單應該修正什麼、為什麼。
紀錄會寫入 changegear_history.db 的 feedback 表，
bot 在後續派單時會自動比對並參考。

執行方式：
    python teach.py                    # 互動：選擇模式（單張 / 批次）
    python teach.py IR-0094069         # 直接指定工單（單張模式）
    python teach.py --list             # 列出所有教學紀錄
    python teach.py --delete <id>      # 刪除指定 feedback 紀錄
    python teach.py --batch            # 直接進入批次貼上模式

批次貼上模式：
    從 Teams/Slack/Email 複製類似訊息：
        IR-0094469 這張放錯 requester，已更正
        IR-0094474 這張不是 error，已修正
        IR-0094471 這張不是 service request: application，已修正為 infrastructure
        IR-0094466 這張是 account management
    每行一筆，bot 會自動抓 ticket ID + 把整行當作教學依據存起來。
"""

import sqlite3
import sys
import os
import re

DB_PATH = "changegear_history.db"

# IR-XXXXXX 樣式（含可選的 IR- 前綴；至少 5 個數字才算）
TICKET_ID_PAT = re.compile(r"\b(IR[\-]?\d{5,})\b", re.IGNORECASE)


# ── DB ──────────────────────────────────────────────────────
def init_feedback_table():
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS feedback (
                id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id           TEXT,
                summary             TEXT,
                description         TEXT,
                requester           TEXT,
                req_item            TEXT,
                wrong_owner         TEXT,
                wrong_assigned_to   TEXT,
                correct_owner       TEXT,
                correct_assigned_to TEXT,
                reason              TEXT,
                created_at          TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()
    finally:
        conn.close()


def get_ticket(ticket_id: str) -> dict | None:
    conn = sqlite3.connect(DB_PATH)
    try:
        r = conn.execute("""
            SELECT ticket_id, summary, description, requester,
                   owner, assigned_to, req_item, inc_parent, inc_child, inc_item
            FROM assignments WHERE ticket_id=?
        """, (ticket_id,)).fetchone()
    finally:
        conn.close()
    if not r:
        return None
    return {
        "ticket_id":   r[0],
        "summary":     r[1] or "",
        "description": r[2] or "",
        "requester":   r[3] or "",
        "owner":       r[4] or "",
        "assigned_to": r[5] or "",
        "req_item":    r[6] or "",
        "inc_parent":  r[7] or "",
        "inc_child":   r[8] or "",
        "inc_item":    r[9] or "",
    }


def save_feedback(t: dict, correct_owner: str, correct_assigned: str, reason: str):
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("""
            INSERT INTO feedback
            (ticket_id, summary, description, requester, req_item,
             wrong_owner, wrong_assigned_to, correct_owner, correct_assigned_to, reason)
            VALUES (?,?,?,?,?,?,?,?,?,?)
        """, (
            t["ticket_id"], t["summary"], t["description"], t["requester"],
            t["req_item"],
            t["owner"], t["assigned_to"],
            correct_owner, correct_assigned,
            reason,
        ))
        conn.commit()
    finally:
        conn.close()


def list_feedback():
    conn = sqlite3.connect(DB_PATH)
    try:
        rows = conn.execute("""
            SELECT id, ticket_id, summary, wrong_assigned_to,
                   correct_assigned_to, reason, created_at
            FROM feedback ORDER BY created_at DESC
        """).fetchall()
    finally:
        conn.close()

    if not rows:
        print("（feedback 表為空）")
        return

    print(f"\n共 {len(rows)} 筆教學紀錄：\n")
    for r in rows:
        fb_id, tid, summary, wrong, correct, reason, created = r
        print(f"  [#{fb_id}] {created}")
        print(f"    工單: {tid}")
        print(f"    摘要: {(summary or '')[:60]}")
        print(f"    ✗ 錯誤: {wrong} → ✓ 正確: {correct}")
        print(f"    原因: {reason}")
        print()


def delete_feedback(fb_id: int):
    conn = sqlite3.connect(DB_PATH)
    try:
        c = conn.execute("DELETE FROM feedback WHERE id=?", (fb_id,))
        conn.commit()
    finally:
        conn.close()
    if c.rowcount > 0:
        print(f"✓ 已刪除 feedback #{fb_id}")
    else:
        print(f"⛔ 找不到 feedback #{fb_id}")


# ── 單張互動模式 ────────────────────────────────────────────
def single_ticket_mode(ticket_id_hint: str = ""):
    ticket_id = ticket_id_hint or input("\n工單 ID（例：IR-0094069）：").strip().upper()
    if not ticket_id:
        print("⛔ 取消")
        return

    t = get_ticket(ticket_id)
    if not t:
        print(f"⛔ 找不到工單 {ticket_id}")
        print("   只有「bot 派過」或「bot 追蹤過」的工單才會在 DB 中")
        return

    print(f"\n找到工單：")
    print(f"  ID:             {t['ticket_id']}")
    print(f"  主旨:           {t['summary']}")
    print(f"  寄件者:         {t['requester']}")
    print(f"  目前 Owner:     {t['owner']}")
    print(f"  目前 Assigned:  {t['assigned_to']}")
    print(f"  Inc Type:       {t['inc_parent']} > {t['inc_child']} > {t['inc_item']}")
    print(f"  Req Item:       {t['req_item']}")

    print("\n──── 請輸入正確的派單資訊（按 Enter 跳過則保留原值）────")
    correct_assigned = input(f"正確的 Assigned To [{t['assigned_to']}]：").strip() or t["assigned_to"]
    correct_owner    = input(f"正確的 Owner       [{t['owner']}]：").strip() or t["owner"]

    print("\n為什麼？這個原因會讓 bot 未來判斷類似工單時參考")
    print("（例如：「Outlook 群組屬於 Exchange 管理員工作」/")
    print("       「這張是 account management 不是 application」）")
    reason = input("原因：").strip()
    if not reason:
        print("⛔ 取消（沒寫原因）")
        return

    print("\n──── 預備寫入 ────")
    print(f"  ✗ 原本派給: Owner={t['owner']} / Assigned={t['assigned_to']}")
    print(f"  ✓ 正確派給: Owner={correct_owner} / Assigned={correct_assigned}")
    print(f"  原因:       {reason}")

    if input("\n確認儲存？(y/N) ").strip().lower() != "y":
        print("已取消")
        return

    save_feedback(t, correct_owner, correct_assigned, reason)
    print(f"\n✅ 已記錄到 feedback 表")


# ── 批次貼上模式（新功能）───────────────────────────────────
def parse_batch_lines(text: str) -> list[dict]:
    """從多行文字中解析出每張工單的修正資訊。

    每行格式（彈性）：
        IR-XXXXXX <任意描述文字>
        IR-XXXXXX： <描述>
        XXXXXX 開頭也接受（缺 IR- 前綴會自動補）

    回傳：[{ticket_id, raw_line, correction_text}, ...]
    """
    items = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        m = TICKET_ID_PAT.search(line)
        if not m:
            continue
        tid = m.group(1).upper()
        # 補上 IR- 前綴並補 0
        if not tid.startswith("IR-"):
            digits = re.sub(r"\D", "", tid)
            tid = f"IR-{digits.zfill(7)}"
        else:
            # 標準化：IR- 後保留所有數字
            digits = re.sub(r"\D", "", tid[3:])
            tid = f"IR-{digits.zfill(7)}"
        # 把 ticket id 之外的部分當作修正描述
        correction = TICKET_ID_PAT.sub("", line).strip(" :,，:、-").strip()
        items.append({
            "ticket_id":       tid,
            "raw_line":        raw.rstrip(),
            "correction_text": correction or "(無描述)",
        })
    return items


def batch_mode():
    print()
    print("──── 批次貼上模式 ────")
    print("貼上多行修正描述（從 Teams/Slack 複製即可）")
    print("每行至少含 IR-XXXXXX 樣式的工單編號 + 任意修正說明")
    print()
    print("輸入結束方式：連續兩次空白 Enter")
    print("    範例：")
    print("        IR-0094469 這張放錯 requester，已更正")
    print("        IR-0094474 這張不是 error，已修正")
    print("        IR-0094471 這張不是 service request: application，已修正為 infrastructure")
    print("        IR-0094466 這張是 account management")
    print()
    print("--- 貼上開始 ---")

    lines = []
    empty_count = 0
    while True:
        try:
            line = input()
        except EOFError:
            break
        if not line.strip():
            empty_count += 1
            if empty_count >= 2:
                break
            continue
        empty_count = 0
        lines.append(line)

    if not lines:
        print("⛔ 沒輸入任何內容，取消")
        return

    items = parse_batch_lines("\n".join(lines))
    if not items:
        print("⛔ 沒解析到任何 IR-XXXXXX 工單編號，取消")
        return

    # ── 對每筆做預覽 ───
    print()
    print("=" * 70)
    print(f"解析到 {len(items)} 筆，預覽：")
    print("=" * 70)
    valid_count = 0
    for i, item in enumerate(items, 1):
        print()
        print(f"[{i}] {item['ticket_id']}")
        t = get_ticket(item["ticket_id"])
        if not t:
            print(f"    ⛔ DB 找不到此工單（bot 沒派過 / 沒追蹤過 / ID 拼錯）")
            print(f"    原文: {item['raw_line']}")
            item["_ticket"] = None
            continue
        item["_ticket"] = t
        valid_count += 1
        print(f"    ✓ 主旨: {t['summary'][:50]}")
        print(f"    目前: Owner={t['owner']} / Assigned={t['assigned_to']}")
        print(f"          Inc={t['inc_parent']} > {t['inc_child']} > {t['inc_item']}")
        print(f"    🎯 教學: {item['correction_text']}")

    print()
    print("=" * 70)
    print(f"摘要：有效 {valid_count} 筆 / 無效 {len(items) - valid_count} 筆")

    if valid_count == 0:
        print("⛔ 沒有任何有效工單可寫入")
        return

    if input(f"\n寫入 {valid_count} 筆教學紀錄？(y/N) ").strip().lower() != "y":
        print("已取消")
        return

    written = 0
    for item in items:
        t = item.get("_ticket")
        if not t:
            continue
        # 批次模式：correct_* 預設等於 wrong_*（描述放在 reason 即可），
        # 讓 Claude 從 reason 文字理解該怎麼修正
        save_feedback(
            t,
            correct_owner=t["owner"],
            correct_assigned=t["assigned_to"],
            reason=item["correction_text"],
        )
        written += 1

    print(f"\n✅ 已寫入 {written} 筆教學紀錄")
    print(f"   bot 下次遇到類似工單會自動參考這些修正描述")


# ── 互動主流程 ──────────────────────────────────────────────
def main():
    if not os.path.exists(DB_PATH):
        print(f"⛔ 找不到 {DB_PATH}")
        print("   請先在同目錄執行過 Auto mission start.bat 至少一次")
        sys.exit(1)

    init_feedback_table()

    args = sys.argv[1:]
    if args and args[0] in ("--list", "-l"):
        list_feedback()
        return
    if args and args[0] in ("--delete", "-d"):
        if len(args) < 2 or not args[1].isdigit():
            print("用法：python teach.py --delete <feedback_id>")
            sys.exit(1)
        delete_feedback(int(args[1]))
        return
    if args and args[0] in ("--batch", "-b"):
        batch_mode()
        return

    # 直接傳入 ticket id → 單張模式
    if args and not args[0].startswith("-"):
        print("=" * 60)
        print("ChangeGear 派單修正回饋工具（單張模式）")
        print("=" * 60)
        single_ticket_mode(args[0].strip().upper())
        return

    # 無參數 → 選擇模式
    print("=" * 60)
    print("ChangeGear 派單修正回饋工具")
    print("=" * 60)
    print()
    print("選擇模式：")
    print("  1. 單張工單   — 完整輸入正確 Owner/Assigned + 原因")
    print("  2. 批次貼上   — 從 Teams/Slack 複製多行修正描述")
    print("  3. 列出紀錄   — 看現有 feedback")
    print("  4. 離開")
    print()

    choice = input("選擇 [1/2/3/4]: ").strip()
    if choice == "1":
        single_ticket_mode()
    elif choice == "2":
        batch_mode()
    elif choice == "3":
        list_feedback()
    else:
        print("再見")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n已中止")
        sys.exit(0)
