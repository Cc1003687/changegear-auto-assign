# ChangeGear Auto-Assign Bot

Automated ticket dispatcher for **ChangeGear ITSM**. Scans new incidents on a schedule and auto-fills assignment fields using a three-layer decision engine:

1. **Claude AI** (Anthropic API) — primary decision layer with structured JSON output
2. **SQLite history DB** — SequenceMatcher fuzzy match against past dispatches
3. **Excel keyword rules** — fallback for rule-based matching
4. **CMDB validation layer** — cross-checks the assignee against the asset's real owner

Originally built for an internal IT team that handles ~50 tickets/day. The bot learns from human corrections and improves over time.

> **Language note**: Most in-code comments and the Excel template are in Traditional Chinese (zh-TW) — that's the language the original team operates in. The code itself, this README, and config keys are documented in English. Translate the comments if you need.

---

## Features

- **Three-layer dispatcher** with weighted trust scoring (manual-corrected × 1.20 > human-dispatched × 1.10 > bot-dispatched × 1.00)
- **CMDB cross-validation**: when confidence is borderline (0.75–0.85), verify the chosen Owner / Assigned-To against the CMDB asset's `team_owner` / `owner`
- **"Hi xxx" name extraction** as fallback when CMDB validation fails — extracts greeting names from email body and re-queries the DB
- **Weekday-based SLA**: Mon–Wed assigns priority 3 with +4 business days; Thu–Sun assigns priority 4 with +5 business days (skipping weekends)
- **Self-learning loop**: every hour the bot re-scans tickets it dispatched, detects human corrections, and updates the DB so the next similar ticket gets routed correctly
- **Popup-based CMDB scraping** via Playwright's `BrowserContext.expect_page()` — necessary because the ChangeGear CMDB detail page only initializes server-side session state when opened via `window.open()`

---

## Architecture

```
┌──────────────────────────────────────────────────────────────────────┐
│                         Auto Assign Bot                              │
│                                                                      │
│   ┌──────────────┐    ┌──────────────┐    ┌──────────────────────┐  │
│   │ APScheduler  │ →  │ Playwright   │ →  │  ChangeGear ITSM     │  │
│   │  every 30min │    │  Async API   │    │  (ASP.NET / IIS)     │  │
│   └──────────────┘    └──────────────┘    └──────────────────────┘  │
│                              │                                       │
│         ┌────────────────────┼────────────────────┐                  │
│         ↓                    ↓                    ↓                  │
│  ┌────────────┐      ┌──────────────┐      ┌──────────────┐         │
│  │ Claude AI  │      │ History DB   │      │ CMDB DB      │         │
│  │ Anthropic  │      │ (SQLite)     │      │ (SQLite)     │         │
│  │ haiku-4-5  │      │ past tickets │      │ asset owners │         │
│  └────────────┘      └──────────────┘      └──────────────┘         │
│                              │                                       │
│                              ↓                                       │
│                       ┌─────────────┐                                │
│                       │ Excel rules │                                │
│                       │  (openpyxl) │                                │
│                       └─────────────┘                                │
└──────────────────────────────────────────────────────────────────────┘
```

---

## Tech Stack

| Layer | Tech |
|---|---|
| Language | Python 3.11+ |
| Browser automation | Playwright async API (Chromium, headless-capable) |
| Popup capture | `BrowserContext.expect_page()` |
| Scheduling | APScheduler (`AsyncIOScheduler`) |
| Storage | SQLite 3 (two DBs: history + CMDB owners) |
| AI | Anthropic Claude API (default `claude-haiku-4-5`) |
| Fuzzy match | Python `difflib.SequenceMatcher` |
| Config | Excel via openpyxl |
| Target system | ChangeGear ITSM (ASP.NET + AngularJS + DevExtreme + Telerik) |

---

## Decision Flow

### AI mode (Claude API key configured)

```
Claude returns dispatch decision with confidence ≥ 0.85
    ↓
confidence > 0.85   → skip CMDB validation, dispatch directly
confidence == 0.85  → run CMDB validation
                        ├─ pass → dispatch
                        └─ fail → extract "Hi xxx" → re-query DB (≥ 0.85) → dispatch
confidence < 0.85   → track only, do not dispatch
```

### Traditional mode (no Claude key)

```
DB SequenceMatcher score
    ↓
score > 0.85        → skip CMDB validation, dispatch directly
0.75 ≤ score ≤ 0.85 → run CMDB validation (same flow as above)
0.65 ≤ score < 0.75 → dispatch directly (low confidence but acceptable)
score < 0.65        → fall through to Excel keyword rules
    ↓
no keyword match    → fall through to defaults from Excel config
```

### Help Desk owner filter

By default the bot only dispatches tickets whose Owner is "Help Desk", to avoid hijacking other teams' tickets. **Exception**: when `_score ≥ 0.75`, this filter is bypassed — the bot trusts the high-confidence match and dispatches regardless of Owner.

---

## CMDB Validation

When confidence falls in the 0.75–0.85 range, the bot cross-validates the dispatch choice against the CMDB asset's real owners:

```
Proposed Assigned To  ←→  CMDB.owner       (asset's individual owner)
Proposed Owner        ←→  CMDB.team_owner  (asset's team owner)
```

Match uses loose comparison (substring contains OR `SequenceMatcher` ratio ≥ 0.75). Both sides must match to pass.

If validation fails, the bot scans the ticket Description for greetings (`Hi xxx`, `Hello xxx`, `Dear xxx`) and does a second DB lookup by name, requiring score ≥ 0.85 to proceed.

---

## Installation

### 1. Clone & install dependencies

```bash
git clone https://github.com/<your-username>/changegear-auto-assign.git
cd changegear-auto-assign
pip install -r requirements.txt
playwright install chromium
```

### 2. Configure

```bash
# Make a working copy of the template
cp ChangeGear_AutoAssign_Rules.example.xlsx ChangeGear_AutoAssign_Rules.xlsx
```

Open `ChangeGear_AutoAssign_Rules.xlsx` and fill in:

| Sheet | What to fill |
|---|---|
| `程式設定` (Settings) | AD account/password, ChangeGear URL, Claude API key (optional), scan interval, similarity threshold |
| `關鍵字派單規則` (Keyword Rules) | Keyword → Owner / Assigned To / Incident Type mapping (replace example rows with your own) |
| `Requester's Item 對應` | Keyword → Requester's Item mapping |

The working copy is gitignored; only the `.example.xlsx` template is tracked.

### 3. (Optional) Customize TARGET_ASSIGNEES in `build_history_db.py`

If you want to broaden DB coverage to specific assignees beyond Help Desk, edit line ~306:

```python
TARGET_ASSIGNEES = ("user_a", "user_b", "user_c")  # ← replace with real AD account substrings
```

These are used for the "Hi xxx" second-match learning. Use lowercase substrings; matching is `in` (containment).

### 4. Build the databases (first-time setup)

```bash
# Scrape historical tickets into changegear_history.db
./Data\ base\ build.bat            # Windows
python build_history_db.py         # Cross-platform

# Scrape CMDB asset owners into cmdb_owners.db
./Build\ CMDB\ DB.bat              # Windows
python build_cmdb_db.py            # Cross-platform
```

Ctrl+C is safe — already-scraped rows are preserved.

### 5. Run the bot

```bash
./Auto\ mission\ start.bat         # Windows
python changegear_auto_assign_v6.py
```

The scheduler scans for new tickets every N minutes (default 30) and runs an hourly correction-learning sweep.

---

## Repository Layout

```
changegear-auto-assign/
├── changegear_auto_assign_v6.py        # Main dispatcher
├── build_history_db.py                 # Historical ticket scraper
├── build_cmdb_db.py                    # CMDB All-Managed-Items scraper
│
├── Auto mission start.bat              # Windows launcher: run dispatcher
├── Data base build.bat                 # Windows launcher: build history DB
├── Build CMDB DB.bat                   # Windows launcher: build CMDB DB
│
├── ChangeGear_AutoAssign_Rules.example.xlsx   # Config template (committed)
├── requirements.txt
├── .gitignore
├── LICENSE                             # MIT
└── README.md
```

The following are **generated at runtime** and gitignored:

| File | Contains |
|---|---|
| `ChangeGear_AutoAssign_Rules.xlsx` | Real credentials + rules |
| `changegear_history.db` | Historical dispatch records |
| `cmdb_owners.db` | CMDB asset owners |
| `auto_assign.log` | Dispatch decisions & errors |
| `build_history.log` | History scraper log |
| `cmdb_build.log` | CMDB scraper log |
| `debug_*.png` | Screenshots of failed dispatch attempts |

---

## Database Schemas

### `changegear_history.db`

```sql
CREATE TABLE assignments (
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
    bot_assigned INTEGER DEFAULT 0,  -- 0=human, 1=bot, 2=tracked-only
    corrected    INTEGER DEFAULT 0,  -- 1=human-corrected after bot dispatch
    created_at   TIMESTAMP
);
```

### `cmdb_owners.db`

```sql
CREATE TABLE cmdb_items (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    oid           TEXT UNIQUE,
    critical_name TEXT,
    item_type     TEXT,
    location      TEXT,
    criticality   TEXT,
    mgmt_status   TEXT,
    department    TEXT,
    op_status     TEXT,
    owner         TEXT,    -- maps to dispatch "Assigned To"
    co_owner      TEXT,
    tech_owner    TEXT,
    team_owner    TEXT,    -- maps to dispatch "Owner"
    users_groups  TEXT,
    created_at    TIMESTAMP,
    updated_at    TIMESTAMP
);
```

---

## Self-Learning Loop

Every hour the bot scans tickets it has dispatched (`bot_assigned = 1`) and tickets it could not dispatch but tracked (`bot_assigned = 2`):

- If a human has since edited the dispatch fields → write the corrected values back to the DB and flag `corrected = 1`
- Tickets with `corrected = 1` receive the highest trust weight (× 1.20) in future similarity matches

Over time, this means the bot's accuracy improves without manual rule tuning.

---

## FAQ

**Q: Why use Playwright popup capture for CMDB?**
The CMDB detail page (`SDItemEditPanel.aspx?boundtable=CIBase`) returns "Object reference not set" on direct navigation because server-side session state is only initialized when the page is opened via the module's `OnGetRowValues(oid)` JS function, which internally calls `window.open()`. Playwright's `BrowserContext.expect_page()` captures that popup cleanly.

**Q: Why two SQLite DBs instead of one?**
History DB grows continuously (one row per dispatched ticket) while CMDB DB is rebuilt periodically from the asset master list. Keeping them separate makes the rebuild cycle independent — you can re-scrape CMDB without losing dispatch history.

**Q: Headless mode hangs on login**
The ChangeGear server uses Windows Integrated Auth (NTLM). The bot passes `http_credentials` to Playwright's `BrowserContext`, which works in both headless and headed modes. If headless fails, verify your AD account isn't locked and check `auto_assign.log` for the actual error.

**Q: Why are some Owner / Assigned To values rejected with "⛔ 二次人名比對信心不足"?**
This means CMDB validation failed AND the secondary "Hi xxx" name lookup couldn't reach 0.85 confidence. The bot intentionally refuses to dispatch in this case — it would rather track the ticket for human review than dispatch incorrectly.

---

## Contributing

PRs welcome, especially:
- Adapters for other ITSM systems (ServiceNow, Jira Service Management, etc.)
- Alternative AI providers (OpenAI, local LLMs via Ollama)
- Tests / sample data fixtures

This was extracted from an internal tool — code quality is "works in production for one team" rather than "polished library". Suggestions for refactoring are appreciated.

---

## License

MIT — see [LICENSE](LICENSE).

> The author has no affiliation with SunView Software / ChangeGear. "ChangeGear" is a trademark of its respective owner. This project is an independent automation tool that interacts with ChangeGear via its standard web UI.
