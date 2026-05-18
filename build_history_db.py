"""
ChangeGear 歷史資料庫建立工具
掃描 ChangeGear 已派單的歷史工單，批次匯入 SQLite 學習資料庫

執行方式:
    python build_history_db.py

功能:
    - 自動分頁爬取 View 的歷史工單（每頁 50 筆）
    - Phase 1 (快速): 從列表抓 summary / owner / assigned_to / ticket_date
    - Phase 2 (深度): 用獨立 page 開啟每張工單取得 incident_type / description（不影響列表頁）
    - 斷點續傳: 已存在的 ticket_id 自動跳過
    - 結束後印出統計報告

儲存條件（OR 任一成立即收錄）:
    1. 工單日期 >= MIN_TICKET_DATE（預設 2024-01-01，「2024 年以後的所有工單」）
    2. Owner 為 Help Desk
    3. Assigned To 為指定人員（TARGET_ASSIGNEES，CMDB 二次人名比對學習用）

跳過條件:
    - 工單日期早於 MIN_TICKET_DATE → 一律跳過（不論 Owner / Assigned To）
"""

import asyncio
import sqlite3
import re
import sys
import logging
from playwright.async_api import async_playwright

# ── 設定 ────────────────────────────────────────────────
DB_PATH  = "changegear_history.db"
BASE_URL = "https://your-changegear-server.example.com/CGWeb"  # TODO: 改成自己公司的 ChangeGear 網址

try:
    import openpyxl
    wb = openpyxl.load_workbook("ChangeGear_AutoAssign_Rules.xlsx")
    ws = wb["程式設定"]
    cfg = {str(r[0]).strip(): str(r[1]).strip() for r in ws.iter_rows(min_row=3, values_only=True) if r[0] and r[1]}
    AD_ACCOUNT  = cfg.get("AD 帳號", "")
    AD_PASSWORD = cfg.get("AD 密碼", "")
    HEADLESS    = cfg.get("headless 模式", "False").lower() == "true"
except Exception:
    AD_ACCOUNT  = ""
    AD_PASSWORD = ""
    HEADLESS    = False

# 要爬取的 View（正確 view 參數名稱來自 AngularJS scope）
VIEWS = [
    ("Incident", "All Incidents"),               # 全量（含 active + closed）
]

# Phase 2 深度爬取每張工單（True=慢但完整；False=快速只取列表欄位）
DEEP_SCRAPE = True
# 最多處理幾張工單（0=無限制）
MAX_TICKETS = 0

# 工單日期下限（ISO 字串格式 YYYY-MM-DD）：早於此日期的工單一律跳過
# 設定為 2024-01-01 代表只爬取 2024 年以後（含）建立／修改的工單
MIN_TICKET_DATE = "2024-01-01"

# 列表 row 內日期匹配 pattern：支援 YYYY-MM-DD、YYYY/MM/DD、M/D/YYYY 三種格式
_DATE_PAT = re.compile(
    r"\b(\d{4})[-/](\d{1,2})[-/](\d{1,2})\b"     # YYYY-MM-DD / YYYY/MM/DD
    r"|\b(\d{1,2})/(\d{1,2})/(\d{4})\b"          # M/D/YYYY（ChangeGear 常見）
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("build_history.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)


# ── SQLite ───────────────────────────────────────────────
def init_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS assignments (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            ticket_id    TEXT UNIQUE,
            summary      TEXT,
            description  TEXT,
            requester    TEXT,
            owner        TEXT,
            assigned_to  TEXT,
            inc_parent   TEXT,
            inc_child    TEXT,
            inc_item     TEXT,
            req_item     TEXT,
            created_at   TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    # 舊版 DB 補欄位（用 PRAGMA 確認欄位存在後再 ALTER，避免 exception 被吞）
    existing_cols = {r[1] for r in conn.execute("PRAGMA table_info(assignments)").fetchall()}
    for col, definition in [
        ("requester", "TEXT DEFAULT ''"),
        ("req_item",  "TEXT DEFAULT ''"),
    ]:
        if col not in existing_cols:
            conn.execute(f"ALTER TABLE assignments ADD COLUMN {col} {definition}")
            log.info(f"DB 欄位補齊: 新增 {col}")
    conn.commit()
    return conn


def db_exists(conn: sqlite3.Connection, ticket_id: str) -> bool:
    row = conn.execute("SELECT 1 FROM assignments WHERE ticket_id=?", (ticket_id,)).fetchone()
    return row is not None


def db_upsert(conn: sqlite3.Connection, data: dict):
    conn.execute("""
        INSERT OR REPLACE INTO assignments
        (ticket_id, summary, description, requester,
         owner, assigned_to, inc_parent, inc_child, inc_item, req_item)
        VALUES (?,?,?,?,?,?,?,?,?,?)
    """, (
        data.get("ticket_id", ""),
        data.get("summary", ""),
        (data.get("description", "") or "")[:500],
        data.get("requester", ""),
        data.get("owner", ""),
        data.get("assigned_to", ""),
        data.get("inc_parent", ""),
        data.get("inc_child", ""),
        data.get("inc_item", ""),
        data.get("req_item", ""),
    ))
    conn.commit()


def db_count(conn: sqlite3.Connection) -> int:
    return conn.execute("SELECT COUNT(*) FROM assignments").fetchone()[0]


# ── Playwright 操作 ──────────────────────────────────────
async def go_to_view(page, module: str, view: str) -> bool:
    """導向指定的 View 列表並等待資料列出現"""
    url = (
        f"{BASE_URL}/MainUI/Common/Modules/BaseModule.aspx"
        f"?ModuleName={module}&view={view.replace(' ', '%20')}&text={view.replace(' ', '%20')}"
    )
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=30000)
        # 等待 grid 資料列出現（最多 20 秒）
        await page.wait_for_selector("tr[id*='DXDataRow']", timeout=20000)
        return True
    except Exception as e:
        log.warning(f"導向 {view} 失敗: {e}")
        return False


async def parse_list_page(page) -> list[dict]:
    """解析當前列表頁，回傳工單基本資料

    欄位索引（All Incidents view 已驗證）：
      _2  = Item ID (label)
      _5  = Requester (.GridDataItem)  ← 寄件者
      _8  = Summary (.GridDataItem)    ← 信件主旨
      _9  = Owner (.GridDataItem)
      _11 = Assign To (.GridDataItem)
    OID 從 row ondblclick: OnGetRowValues('<OID>') 取得
    """
    tickets = []
    rows = await page.query_selector_all("tr[id*='DXDataRow']")
    for row in rows:
        try:
            ondblclick = await row.get_attribute("ondblclick") or ""
            m = re.search(r"OnGetRowValues\('(\d+)'\)", ondblclick)
            oid = m.group(1) if m else None

            item_el  = await row.query_selector("td[id$='_2'] label")
            item_id  = (await item_el.inner_text()).strip() if item_el else ""

            req_el   = await row.query_selector("td[id$='_5'] .GridDataItem")
            requester = (await req_el.inner_text()).strip() if req_el else ""

            sum_el   = await row.query_selector("td[id$='_8'] .GridDataItem")
            summary  = (await sum_el.inner_text()).strip() if sum_el else ""

            own_el   = await row.query_selector("td[id$='_9'] .GridDataItem")
            owner    = (await own_el.inner_text()).strip() if own_el else ""

            asgn_el  = await row.query_selector("td[id$='_11'] .GridDataItem")
            assigned = (await asgn_el.inner_text()).strip() if asgn_el else ""

            # ── 掃描整列所有 cell，萃取最新的日期字串（不依賴特定欄位索引）──
            #    支援 ChangeGear 常見的 M/D/YYYY 與 YYYY-MM-DD 格式
            ticket_date = ""
            cells = await row.query_selector_all(".GridDataItem")
            for cell in cells:
                txt = (await cell.inner_text() or "").strip()
                dm  = _DATE_PAT.search(txt)
                if not dm:
                    continue
                if dm.group(1):                                # YYYY-MM-DD
                    y, mo, d = dm.group(1), dm.group(2), dm.group(3)
                else:                                          # M/D/YYYY
                    mo, d, y = dm.group(4), dm.group(5), dm.group(6)
                iso = f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"
                if iso > ticket_date:
                    ticket_date = iso   # 取列內最新的日期作為「Modified Date」近似值

            if item_id and (owner or assigned):
                tickets.append({
                    "oid": oid,
                    "ticket_id": item_id,
                    "requester": requester,
                    "summary": summary,
                    "owner": owner,
                    "assigned_to": assigned,
                    "ticket_date": ticket_date,
                })
        except Exception:
            continue
    return tickets


async def get_next_page(page) -> bool:
    """點擊下一頁，回傳是否成功

    DevExpress pager onclick 規則：
      'PBN' = Next（下一頁）
      'PBP' = Prev（上一頁）
    用 onclick*='PBN' 精確鎖定，避免在第2頁以後誤抓上一頁按鈕。
    """
    try:
        # 找 onclick 含 'PBN' 且未被 disable 的下一頁按鈕
        next_btn = await page.query_selector(
            "a[onclick*=\"'PBN'\"]:not(.dxp-disabledButton)"
        )
        if not next_btn:
            return False
        await next_btn.click()
        await page.wait_for_selector("tr[id*='DXDataRow']", timeout=15000)
        return True
    except Exception:
        return False


_RIRC = "DynamicLayoutControl1_ImpactedResourcesPIT_irc_18"

async def scrape_ticket_detail(detail_page, oid: str) -> dict:
    """用獨立 page 開啟工單詳細頁，取得 incident_type / description / req_item
    不影響列表頁的當前頁碼。

    已驗證的 selector：
      Incident Type  : [id*='IncidentRequestType'][id$='_I'] → input.value
      Description    : div[id*='Description'] .dxrte-content 或 textarea[id*='Description']
      Requester Item : #{_RIRC}_dropDownTextBox_I → input.value (多項逗號分隔)
                       備援: #{_RIRC}_ASPxPopupControl1_SelectedItemsLB → SELECT options
    """
    result = {"inc_parent": "", "inc_child": "", "inc_item": "", "description": "", "req_item": ""}
    try:
        url = (
            f"{BASE_URL}/MainUI/ServiceDesk/SDItemEditPanel.aspx"
            f"?boundtable=IIncidentRequest&CloseOnPerformAction=false"
            f"&ID={oid}&windowWidth=1050&refreshOnClose=true"
        )
        await detail_page.goto(url, wait_until="domcontentloaded", timeout=25000)
        await detail_page.wait_for_load_state("networkidle", timeout=10000)

        # Incident Type 文字輸入框
        inc_input = await detail_page.query_selector("[id*='IncidentRequestType'][id$='_I']")
        if inc_input:
            inc_text = (await inc_input.input_value()).strip()
            parts = [p.strip() for p in re.split(r"[-:>]", inc_text) if p.strip()]
            if len(parts) >= 3:
                result["inc_parent"] = parts[0]
                result["inc_child"]  = parts[1]
                result["inc_item"]   = parts[2]
            elif len(parts) == 2:
                result["inc_parent"] = parts[0]
                result["inc_child"]  = parts[1]

        # Description — 在 divComment 區塊（已驗證 selector）
        desc_el = await detail_page.query_selector(
            "[id*='Description'][id*='_divComment']"
        )
        if desc_el:
            result["description"] = (await desc_el.inner_text()).strip()[:500]

        # Requester's Item — 主 selector: dropDownTextBox_I（input.value，多項逗號分隔）
        ri_input = await detail_page.query_selector(f"#{_RIRC}_dropDownTextBox_I")
        if ri_input:
            ri_val = (await ri_input.input_value()).strip()
            if ri_val:
                result["req_item"] = ri_val
        # 備援: SelectedItemsLB SELECT options
        if not result["req_item"]:
            ri_lb = await detail_page.query_selector(f"#{_RIRC}_ASPxPopupControl1_SelectedItemsLB")
            if ri_lb:
                options = await ri_lb.query_selector_all("option")
                texts = []
                for opt in options:
                    t = (await opt.inner_text()).strip()
                    if t:
                        texts.append(t)
                result["req_item"] = "; ".join(texts)

    except Exception as e:
        log.debug(f"深度爬取失敗 Oid={oid}: {e}")
    return result


# ── 主流程 ───────────────────────────────────────────────
async def main():
    conn = init_db()
    start_count = db_count(conn)
    log.info(f"DB 初始記錄數: {start_count}")

    total_seen = total_saved = total_skipped = 0

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=HEADLESS)
        context = await browser.new_context(
            http_credentials={"username": AD_ACCOUNT, "password": AD_PASSWORD}
        )
        list_page   = await context.new_page()
        detail_page = await context.new_page()  # 獨立 page，不影響列表頁碼

        for module, view in VIEWS:
            log.info(f"══ 開始爬取 View: [{view}] ══")
            if not await go_to_view(list_page, module, view):
                continue

            page_num = 1
            while True:
                log.info(f"  第 {page_num} 頁...")
                tickets = await parse_list_page(list_page)
                log.info(f"  本頁找到 {len(tickets)} 筆有效工單")

                for t in tickets:
                    total_seen += 1

                    # ── 日期過濾：早於 MIN_TICKET_DATE 一律跳過 ───────────
                    ticket_date = t.get("ticket_date") or ""
                    if ticket_date and ticket_date < MIN_TICKET_DATE:
                        total_skipped += 1
                        continue

                    # ── 儲存條件（OR 邏輯，任一條件成立即收錄）─────────────
                    #   1. ticket_date >= MIN_TICKET_DATE（2024 年以後的所有 ticket）
                    #   2. Owner 為 Help Desk
                    #   3. Assigned To 為指定人員（CMDB 二次人名比對學習用）
                    # 範例值請依實際組織人員 AD 帳號替換（小寫，substring 比對即可命中）
                    TARGET_ASSIGNEES = ("user_a", "user_b", "user_c")
                    owner_val    = (t.get("owner") or "").strip().lower()
                    assigned_val = (t.get("assigned_to") or "").strip().lower()
                    is_recent    = bool(ticket_date) and ticket_date >= MIN_TICKET_DATE
                    is_helpdesk  = ("help desk" in owner_val) or ("helpdesk" in owner_val)
                    is_target_assignee = any(
                        target in assigned_val for target in TARGET_ASSIGNEES
                    )
                    if not (is_recent or is_helpdesk or is_target_assignee):
                        total_skipped += 1
                        continue

                    if db_exists(conn, t["ticket_id"]):
                        total_skipped += 1
                        continue

                    if DEEP_SCRAPE and t.get("oid"):
                        detail = await scrape_ticket_detail(detail_page, t["oid"])
                        t.update(detail)

                    db_upsert(conn, t)
                    total_saved += 1

                    if total_saved % 50 == 0:
                        log.info(f"  已儲存 {total_saved} 筆...")

                    if MAX_TICKETS > 0 and total_saved >= MAX_TICKETS:
                        log.info(f"已達上限 {MAX_TICKETS} 筆，停止")
                        break

                if MAX_TICKETS > 0 and total_saved >= MAX_TICKETS:
                    break

                has_next = await get_next_page(list_page)
                if not has_next:
                    log.info(f"  [{view}] 已到最後一頁（共 {page_num} 頁）")
                    break
                page_num += 1

        await browser.close()

    end_count = db_count(conn)
    conn.close()

    log.info("══════════════════════════════")
    log.info(
        f"爬取完成！（儲存條件：ticket_date >= {MIN_TICKET_DATE} "
        f"或 Owner=Help Desk 或 Assigned To ∈ TARGET_ASSIGNEES）"
    )
    log.info(f"  掃描工單數 : {total_seen}")
    log.info(f"  新增記錄數 : {total_saved}")
    log.info(f"  跳過（日期過早 / 不符條件 / 已存在）: {total_skipped}")
    log.info(f"  DB 總記錄數: {end_count}（原 {start_count}）")
    log.info("══════════════════════════════")


if __name__ == "__main__":
    print("=" * 50)
    print("ChangeGear 歷史資料庫建立工具")
    print(f"  DB 路徑    : {DB_PATH}")
    print(f"  深度爬取   : {DEEP_SCRAPE}")
    print(f"  最大筆數   : {'無限制' if MAX_TICKETS == 0 else MAX_TICKETS}")
    print(f"  日期下限   : {MIN_TICKET_DATE}（早於此日期一律跳過）")
    print(f"  headless   : {HEADLESS}")
    print("=" * 50)
    print("按 Enter 開始，Ctrl+C 隨時中止...")
    try:
        input()
    except KeyboardInterrupt:
        sys.exit(0)

    asyncio.run(main())
