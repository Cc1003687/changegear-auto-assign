"""
ChangeGear CMDB All Managed Items 爬取工具
=============================================
從 CMDB → All Managed Items 列表爬取所有項目，
進入每個 item 的 Usage tab 取得 Owner / Co-Owner / Technical Owner / Team Owner，
全部儲存到 SQLite DB（cmdb_owners.db）。

執行方式:
    python build_cmdb_db.py

功能:
    - Phase 1（快速）: 列表頁抓 Critical Name / Type / Location / OID
    - Phase 2（深度）: 透過 OnGetRowValues(oid) 開啟 popup 視窗，
                       讀 General tab + 切換 Usage tab，讀取所有 owner 欄位
    - 斷點續傳: 已存在的 OID 自動跳過
    - 逐頁處理: 每頁解析後立即爬取 Usage 資料，再換下一頁

已驗證的 selector（透過瀏覽器 DOM 直接確認）：
  ─ General Tab (DynamicLayoutControl1_ctl01_C0) ─
    Criticality   : SELECT[id*='ResourceCriticality_ddl']   → selected text
    Mgmt Status   : SELECT[id*='ManagementStatus_ec'][id*='EntityDropDownList'] → text
    Department    : INPUT[id*='Department_tvdd'][id$='_I']   → value
    Location      : INPUT[id*='Location_tvdd'][id$='_I']    → value
    Op. Status    : SELECT[id*='Status_ddl']                → text

  ─ Usage Tab (DynamicLayoutControl1_ctl01_C5) ─
    Usage tab link: A#DynamicLayoutControl1_ctl01_T5T
    Owner         : INPUT[id*='Owner_PersonChooser_GridLookupPC_I'] （排除 CoOwner/TechnicalOwner）
    Co-Owner      : INPUT[id*='CoOwner_PersonChooser_GridLookupPC_I']
    Tech Owner    : INPUT[id*='TechnicalOwner_PersonChooser_GridLookupPC_I']
    Team Owner    : SELECT[id*='ItemOwner_ddl']
    Users/Groups  : TEXTAREA[id*='txtImactedUsers']
"""

import asyncio
import sqlite3
import re
import sys
import logging
from playwright.async_api import async_playwright, Page, BrowserContext

# ── 設定 ─────────────────────────────────────────────────────────────
DB_PATH  = "cmdb_owners.db"
BASE_URL = "https://your-changegear-server.example.com/CGWeb"  # TODO: 改成自己公司的 ChangeGear 網址
CMDB_VIEW_URL = (
    f"{BASE_URL}/MainUI/Common/Modules/BaseModule.aspx"
    "?ModuleName=CMDB&view=All%20Managed%20Items&text=All%20Managed%20Items"
)

try:
    import openpyxl
    wb = openpyxl.load_workbook("ChangeGear_AutoAssign_Rules.xlsx")
    ws = wb["程式設定"]
    cfg = {str(r[0]).strip(): str(r[1]).strip()
           for r in ws.iter_rows(min_row=3, values_only=True) if r[0] and r[1]}
    AD_ACCOUNT  = cfg.get("AD 帳號", "")
    AD_PASSWORD = cfg.get("AD 密碼", "")
    HEADLESS    = cfg.get("headless 模式", "False").lower() == "true"
except Exception:
    AD_ACCOUNT = AD_PASSWORD = ""
    HEADLESS   = False

# 0 = 無限制（爬取全部）
MAX_ITEMS = 0

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("cmdb_build.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger(__name__)


# ── SQLite ────────────────────────────────────────────────────────────
def init_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS cmdb_items (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            oid             TEXT UNIQUE,
            critical_name   TEXT,
            item_type       TEXT,
            location        TEXT,
            criticality     TEXT,
            mgmt_status     TEXT,
            department      TEXT,
            op_status       TEXT,
            owner           TEXT,
            co_owner        TEXT,
            tech_owner      TEXT,
            team_owner      TEXT,
            users_groups    TEXT,
            created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    return conn


def db_exists(conn: sqlite3.Connection, oid: str) -> bool:
    return conn.execute(
        "SELECT 1 FROM cmdb_items WHERE oid=?", (oid,)
    ).fetchone() is not None


def db_upsert(conn: sqlite3.Connection, data: dict):
    conn.execute("""
        INSERT OR REPLACE INTO cmdb_items
        (oid, critical_name, item_type, location,
         criticality, mgmt_status, department, op_status,
         owner, co_owner, tech_owner, team_owner, users_groups,
         updated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?, CURRENT_TIMESTAMP)
    """, (
        data.get("oid", ""),
        data.get("critical_name", ""),
        data.get("item_type", ""),
        data.get("location", ""),
        data.get("criticality", ""),
        data.get("mgmt_status", ""),
        data.get("department", ""),
        data.get("op_status", ""),
        data.get("owner", ""),
        data.get("co_owner", ""),
        data.get("tech_owner", ""),
        data.get("team_owner", ""),
        data.get("users_groups", "")[:300] if data.get("users_groups") else "",
    ))
    conn.commit()


def db_count(conn: sqlite3.Connection) -> int:
    return conn.execute("SELECT COUNT(*) FROM cmdb_items").fetchone()[0]


# ── 列表頁解析 ────────────────────────────────────────────────────────
async def parse_list_page(page: Page) -> list[dict]:
    """
    解析 All Managed Items 列表頁。
    已驗證欄位 index（All Managed Items view）：
      _3  = Critical Name (.GridDataItem)
      _4  = Type         (.GridDataItem)
      _7  = Location     (.GridDataItem)
    OID 從 ondblclick: OnGetRowValues('<OID>') 取得
    """
    items = []
    rows = await page.query_selector_all("tr[id*='DXDataRow']")
    for row in rows:
        try:
            ondblclick = await row.get_attribute("ondblclick") or ""
            m = re.search(r"OnGetRowValues\('(\d+)'\)", ondblclick)
            if not m:
                continue
            oid = m.group(1)

            name_el = await row.query_selector("td[id$='_3'] .GridDataItem")
            critical_name = (await name_el.inner_text()).strip() if name_el else ""

            type_el = await row.query_selector("td[id$='_4'] .GridDataItem")
            item_type = (await type_el.inner_text()).strip() if type_el else ""

            loc_el = await row.query_selector("td[id$='_7'] .GridDataItem")
            location = (await loc_el.inner_text()).strip() if loc_el else ""

            if oid:
                items.append({
                    "oid": oid,
                    "critical_name": critical_name,
                    "item_type": item_type,
                    "location": location,
                })
        except Exception:
            continue
    return items


async def get_next_page(page: Page) -> bool:
    """點擊下一頁，同 build_history_db.py 的 PBN 邏輯。"""
    try:
        btn = await page.query_selector(
            "a[onclick*=\"'PBN'\"]:not(.dxp-disabledButton)"
        )
        if not btn:
            return False
        await btn.click()
        await page.wait_for_selector("tr[id*='DXDataRow']", timeout=15000)
        return True
    except Exception:
        return False


# ── Detail 頁解析（透過 popup）────────────────────────────────────────
async def scrape_item_via_popup(list_page: Page, context: BrowserContext,
                                item: dict) -> dict:
    """
    透過 list_page 呼叫 OnGetRowValues(oid)，以 context.expect_page() 捕捉
    彈出的 popup 視窗，讀取 General tab 與 Usage tab 欄位。

    注意：OnGetRowValues() 是 ChangeGear CMDB module 注入的全域 JS 函數，
    直接以 window.open() 開啟 SDItemEditPanel.aspx?boundtable=CIBase&ID=...
    伺服器端需要由此方式開啟（直接 navigate 會因 session state 缺失而報錯）。
    """
    oid = item["oid"]
    detail = {
        "criticality": "", "mgmt_status": "", "department": "",
        "op_status": "", "owner": "", "co_owner": "",
        "tech_owner": "", "team_owner": "", "users_groups": "",
    }
    popup = None

    try:
        async with context.expect_page(timeout=12000) as page_info:
            await list_page.evaluate(f"OnGetRowValues('{oid}')")

        popup = await page_info.value
        await popup.wait_for_load_state("domcontentloaded", timeout=20000)
        await popup.wait_for_load_state("networkidle", timeout=8000)

        # ── General Tab（預設就是 C0，直接讀）────────────────────────
        async def get_input(sel: str) -> str:
            try:
                el = await popup.query_selector(sel)
                if not el:
                    return ""
                tag = await el.evaluate("el => el.tagName")
                if tag == "SELECT":
                    return await el.evaluate(
                        "el => el.options[el.selectedIndex]?.text?.trim() || ''"
                    )
                return (await el.input_value()).strip()
            except Exception:
                return ""

        detail["criticality"]  = await get_input("select[id*='ResourceCriticality_ddl']")
        detail["mgmt_status"]  = await get_input(
            "select[id*='ManagementStatus_ec'][id*='EntityDropDownList']"
        )
        detail["department"]   = await get_input("[id*='Department_tvdd'][id$='_I']")
        detail["op_status"]    = await get_input("select[id*='Status_ddl']")

        # ── 點擊 Usage Tab（T5T）→ 等待 C5 顯示 ─────────────────────
        usage_tab_link = await popup.query_selector(
            "#DynamicLayoutControl1_ctl01_T5T"
        )
        if usage_tab_link:
            await usage_tab_link.click()
            # 等待 C5 panel 激活（DevExtreme 切換方式：移除 display:none inline style）
            # C5 激活後 style.display = ""（空，由 CSS class 控制顯示）
            # C0 被隱藏時 style.display = "none"
            try:
                await popup.wait_for_function(
                    """() => {
                        var c0 = document.getElementById('DynamicLayoutControl1_ctl01_C0');
                        return c0 && c0.style.display === 'none';
                    }""",
                    timeout=8000
                )
            except Exception:
                await popup.wait_for_timeout(1500)
        else:
            log.debug(f"OID={oid}: Usage tab link 未找到，嘗試備援 selector")
            backup = await popup.query_selector("a:text('Usage')")
            if backup:
                await backup.click()
                await popup.wait_for_timeout(1500)

        # ── Usage Tab 欄位讀取 ────────────────────────────────────────
        # Owner（Person Chooser，排除 CoOwner / TechnicalOwner）
        detail["owner"] = await get_input(
            "[id$='Owner_PersonChooser_GridLookupPC_I']:not([id*='CoOwner']):not([id*='Technical'])"
        )
        detail["co_owner"]    = await get_input("[id*='CoOwner_PersonChooser_GridLookupPC_I']")
        detail["tech_owner"]  = await get_input("[id*='TechnicalOwner_PersonChooser_GridLookupPC_I']")
        detail["team_owner"]  = await get_input("select[id*='ItemOwner_ddl']")

        # users_groups（TEXTAREA id 中含 txtImactedUsers，拼字故意如此）
        ug_el = await popup.query_selector("textarea[id*='txtImactedUsers']")
        if ug_el:
            detail["users_groups"] = (await ug_el.text_content() or "").strip()[:300]

        await popup.close()

    except Exception as e:
        log.debug(f"popup OID={oid} 失敗: {e}")
        if popup is not None:
            try:
                await popup.close()
            except Exception:
                pass

    return {**item, **detail}


# ── 主流程 ────────────────────────────────────────────────────────────
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
        list_page = await context.new_page()

        # ── 導向 All Managed Items ────────────────────────────────────
        log.info("載入 CMDB All Managed Items...")
        await list_page.goto(CMDB_VIEW_URL, wait_until="domcontentloaded", timeout=30000)
        try:
            await list_page.wait_for_selector("tr[id*='DXDataRow']", timeout=20000)
        except Exception:
            log.error("列表頁載入失敗，請確認帳號密碼與網址是否正確")
            await browser.close()
            return

        log.info("列表載入成功，開始逐頁爬取...")

        # ── 逐頁處理 ─────────────────────────────────────────────────
        page_num = 1
        while True:
            log.info(f"  第 {page_num} 頁...")
            items = await parse_list_page(list_page)
            log.info(f"  本頁找到 {len(items)} 筆")

            for item in items:
                total_seen += 1

                if db_exists(conn, item["oid"]):
                    total_skipped += 1
                    continue

                # 透過 popup 深度爬取
                full_data = await scrape_item_via_popup(list_page, context, item)
                db_upsert(conn, full_data)
                total_saved += 1

                log.info(
                    f"  [{total_saved:4d}] {full_data['critical_name'][:35]:<35} "
                    f"| Owner={full_data.get('owner','')[:20]:<20} "
                    f"| Team={full_data.get('team_owner','')}"
                )

                if total_saved % 50 == 0:
                    log.info(f"  ── 已儲存 {total_saved} 筆 ──")

                if MAX_ITEMS > 0 and total_saved >= MAX_ITEMS:
                    log.info(f"已達上限 {MAX_ITEMS} 筆，停止")
                    break

            if MAX_ITEMS > 0 and total_saved >= MAX_ITEMS:
                break

            has_next = await get_next_page(list_page)
            if not has_next:
                log.info(f"  已到最後一頁（共 {page_num} 頁）")
                break
            page_num += 1

        await browser.close()

    end_count = db_count(conn)
    conn.close()

    log.info("══════════════════════════════════════════")
    log.info(f"爬取完成！")
    log.info(f"  掃描項目數  : {total_seen}")
    log.info(f"  新增記錄數  : {total_saved}")
    log.info(f"  跳過（已存）: {total_skipped}")
    log.info(f"  DB 總記錄數 : {end_count}（原 {start_count}）")
    log.info("══════════════════════════════════════════")


if __name__ == "__main__":
    print("=" * 60)
    print("ChangeGear CMDB Owner 資料庫建立工具")
    print(f"  DB 路徑  : {DB_PATH}")
    print(f"  目標     : All Managed Items（810 筆預估）")
    print(f"  headless : {HEADLESS}")
    print("=" * 60)
    print("按 Enter 開始，Ctrl+C 隨時中止（已存資料不會遺失）...")
    try:
        input()
    except KeyboardInterrupt:
        sys.exit(0)

    asyncio.run(main())
