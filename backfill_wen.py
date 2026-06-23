"""
Wen 驗證 Backfill 工具
======================

重抳 ChangeGear 每張工單的審計記錄，把最後 Save/Accept 者填進
assignments.last_saver 與 assignments.is_wen_blessed。

之後 bot 在 db_find_similar / db_get_candidates 會自動把 wen.hsieh
驗證過的歷史單以 ×2 權重列入優先參考。

執行方式：
    python backfill_wen.py              # 全部跑（跳過已填過的）
    python backfill_wen.py --force      # 全部重抳（即使 last_saver 已填）
    python backfill_wen.py --limit 100  # 只跑前 100 張（測試用）

中止安全：Ctrl+C 隨時中止，已寫的不會掉。再跑一次會接續上次未處理的。
"""

import asyncio
import sqlite3
import sys
import os

# 共用 changegear_auto_assign_v6 內的設定與 helper
import changegear_auto_assign_v6 as m

DB_PATH = m.DB_PATH


async def main():
    args = sys.argv[1:]
    force = "--force" in args
    limit = None
    if "--limit" in args:
        idx = args.index("--limit")
        if idx + 1 < len(args):
            limit = int(args[idx + 1])

    print("=" * 60)
    print(f"  Wen Backfill — force={force}, limit={limit or '無上限'}")
    print("=" * 60)

    # 1. 取出需要 backfill 的工單清單
    conn = sqlite3.connect(DB_PATH)
    where = "" if force else "WHERE (last_saver IS NULL OR last_saver = '')"
    rows = conn.execute(f"""
        SELECT ticket_id, oid FROM assignments
        {where}
        ORDER BY created_at DESC
        {f'LIMIT {limit}' if limit else ''}
    """).fetchall()
    conn.close()

    total = len(rows)
    if total == 0:
        print("\n沒有需要 backfill 的工單（全部已有 last_saver，可加 --force 強制重抳）")
        return
    print(f"\n預計處理 {total} 張工單")
    print("Ctrl+C 隨時中止")
    print()

    # 2. 啟動 bot 框架（重用瀏覽器、登入流程）
    bot = m.ChangeGearBot()
    await bot.start()
    print("瀏覽器已啟動，開始 backfill...\n")

    processed = wen_count = unknown_count = 0
    try:
        for i, (ticket_id, oid) in enumerate(rows, 1):
            if not oid:
                m.log.debug(f"{ticket_id}: 無 oid，跳過")
                continue
            try:
                url = (
                    f"{m.RULES['base_url']}/MainUI/ServiceDesk/SDItemEditPanel.aspx"
                    f"?boundtable=IIncidentRequest&CloseOnPerformAction=false"
                    f"&ID={oid}&windowWidth=1050&refreshOnClose=true"
                )
                await bot.page.goto(url, wait_until="domcontentloaded", timeout=20000)
                await bot.page.wait_for_timeout(1200)

                last_saver = await bot.read_last_saver(bot.page)
                wen = 1 if bot._is_wen(last_saver) else 0

                conn = sqlite3.connect(DB_PATH)
                conn.execute(
                    "UPDATE assignments SET last_saver=?, is_wen_blessed=? WHERE ticket_id=?",
                    (last_saver or "", wen, ticket_id),
                )
                conn.commit()
                conn.close()

                processed += 1
                if wen:
                    wen_count += 1
                elif not last_saver:
                    unknown_count += 1

                tag = "🌟Wen" if wen else (last_saver or "?")
                print(f"  [{i:4d}/{total}] {ticket_id}  →  {tag}")

            except Exception as e:
                m.log.warning(f"{ticket_id} backfill 失敗: {e}")
                continue
    finally:
        await bot.stop()

    print()
    print("=" * 60)
    print(f"完成！")
    print(f"  已處理:        {processed} / {total}")
    print(f"  Wen 驗證過:    {wen_count} 張  ({wen_count/processed*100:.1f}%)" if processed else "")
    print(f"  Saver 抓不到:  {unknown_count} 張（審計區塊可能用不同 selector）")
    print(f"\n下次 bot 派單時，這些 Wen 驗證過的歷史單會自動 ×2 權重優先參考")
    print("=" * 60)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n已中止（已寫的資料不會掉）")
        sys.exit(0)
