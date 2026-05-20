"""
ChangeGear 派單修正回饋工具
============================

互動式工具：告訴 bot 哪張工單應該派給誰、為什麼。
紀錄會寫入 changegear_history.db 的 feedback 表，
bot 在後續派單時會自動比對並參考。

執行方式：
    python teach.py                    # 互動：手動輸入工單 ID
    python teach.py IR-0094069         # 指定工單 ID
    python teach.py --list             # 列出所有教學紀錄
    python teach.py --delete <id>      # 刪除指定 feedback 紀錄

工作流程：
    1. bot 派錯一張工單
    2. 你在 ChangeGear UI 手動修正派單對象
    3. 跑 teach.py，輸入該工單 ID
    4. 系統列出工單摘要 + 原本派錯的對象
    5. 你輸入「正確應派給誰」+「為什麼」（讓 bot 學會判斷依據）
    6. 寫入後，後續類似工單會被 bot 主動參考此教訓
"""

import sqlite3
import sys
import os

DB_PATH = "changegear_history.db"


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
    """從 assignments 表取得工單目前的派單資訊（bot 派過或追蹤過才有）。"""
    conn = sqlite3.connect(DB_PATH)
    try:
        r = conn.execute("""
            SELECT ticket_id, summary, description, requester,
                   owner, assigned_to, req_item
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


# ── 互動主流程 ──────────────────────────────────────────────
def main():
    # 確認 DB 存在
    if not os.path.exists(DB_PATH):
        print(f"⛔ 找不到 {DB_PATH}")
        print("   請先在同目錄執行過 Auto mission start.bat 至少一次")
        sys.exit(1)

    init_feedback_table()

    # ── 解析參數 ──
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

    print("=" * 60)
    print("ChangeGear 派單修正回饋工具")
    print("告訴 bot 哪張工單應該派給誰、為什麼")
    print("=" * 60)

    ticket_id = args[0].strip().upper() if args else input("\n工單 ID（例：IR-0094069）：").strip().upper()
    if not ticket_id:
        print("⛔ 取消")
        sys.exit(0)

    t = get_ticket(ticket_id)
    if not t:
        print(f"⛔ 找不到工單 {ticket_id}")
        print("   只有「bot 派過」或「bot 追蹤過」的工單才會在 DB 中")
        sys.exit(1)

    print(f"\n找到工單：")
    print(f"  ID:             {t['ticket_id']}")
    print(f"  主旨:           {t['summary']}")
    print(f"  寄件者:         {t['requester']}")
    print(f"  目前 Owner:     {t['owner']}")
    print(f"  目前 Assigned:  {t['assigned_to']}")
    print(f"  Req Item:       {t['req_item']}")

    print("\n──── 請輸入正確的派單資訊 ────")
    correct_assigned = input("正確的 Assigned To（AD 帳號或顯示名稱）：").strip()
    if not correct_assigned:
        print("⛔ 取消（沒填 Assigned To）")
        sys.exit(0)

    correct_owner = input("正確的 Owner（Enter 採用 'Help Desk'）：").strip() or "Help Desk"

    print("\n為什麼？這個原因會讓 bot 未來判斷類似工單時參考")
    print("（例如：「Outlook 群組屬於 Exchange 管理員工作，不該派給一般 Help Desk」）")
    reason = input("原因：").strip()
    if not reason:
        print("⛔ 取消（沒寫原因）")
        sys.exit(0)

    print("\n──── 預備寫入 ────")
    print(f"  ✗ 原本派給: Owner={t['owner']} / Assigned={t['assigned_to']}")
    print(f"  ✓ 正確派給: Owner={correct_owner} / Assigned={correct_assigned}")
    print(f"  原因:       {reason}")

    confirm = input("\n確認儲存？(y/N) ").strip().lower()
    if confirm != "y":
        print("已取消")
        sys.exit(0)

    save_feedback(t, correct_owner, correct_assigned, reason)
    print(f"\n✅ 已記錄到 feedback 表")
    print(f"   下次有類似工單，bot 會自動參考此教訓並優先採用正確派單對象")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n已中止")
        sys.exit(0)
