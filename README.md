# ChangeGear Auto-Assign Bot

Automated ticket dispatcher for **ChangeGear ITSM**. Scans new incidents on a schedule and auto-fills the assignment fields using a **four-layer learning engine** that combines:

1. **Human feedback** (`teach.py`) — explicit, reasoned corrections from the operator
2. **Claude AI** (Anthropic API) — primary arbiter with three-signal context
3. **CMDB** (`cmdb_owners.db`) — asset-ownership ground truth
4. **History DB** (`changegear_history.db`) — SequenceMatcher fuzzy match against past dispatches

Plus an **Excel keyword fallback** for rule-based matching and self-learning loops for both implicit corrections (someone edits a bot-dispatched ticket in the UI) and explicit corrections (operator runs `teach.py`).

Originally built for an internal IT team that handles ~50 tickets/day. The bot becomes more accurate every week as human feedback accumulates.

> **Language note:** In-code comments and the Excel template are in Traditional Chinese (zh-TW) — the language the original team operates in. The code itself, this README, and config keys are documented in English. Translate the comments if you need.

---

## Features

- **Four-layer dispatcher** with explicit weight ordering
  (feedback > AI > CMDB > history > Excel > default)
- **Three-signal unified decision** — every Claude call sees ticket content, top-N historical candidates, AND each candidate's CMDB ownership note in a single prompt. No multi-stage fallback dance; the model arbitrates once with full context.
- **CMDB as a first-class dispatch source**, not just a validator. When a ticket maps cleanly to a known CI, the asset's owner becomes the dispatch target directly.
- **Operator feedback loop** (`teach.py`): tell the bot "this ticket should have gone to X because Y." The reason becomes a lesson injected into future Claude prompts.
- **Weekday-based SLA**: Mon–Wed → priority 3 with +4 business days; Thu–Sun → priority 4 with +5 business days (skipping weekends).
- **Self-learning loop**: every hour the bot re-scans tickets it dispatched and tickets it tracked-but-didn't-dispatch, detects human corrections, and upgrades the DB so the next similar ticket is routed correctly.
- **Confidence-gated Help Desk filter**: only dispatches non-Help-Desk tickets when the dispatch confidence is ≥ 0.75 — high-confidence matches bypass the filter; low-confidence ones are tracked-only.
- **Prompt caching enabled** — stable parts of the prompt (rules, CI list, Incident Type list, JSON format) are cached at ~0.1× cost; only volatile per-ticket content is sent uncached.
- **Popup-based CMDB scraping** via Playwright's `BrowserContext.expect_page()` — necessary because the ChangeGear CMDB detail page only initializes server-side session state when opened via `window.open()`.
- **Person Chooser AD-account extraction** — strips display strings like `"wen.hsieh 謝文譯"` down to `"wen.hsieh"` before typing into ChangeGear's autocomplete, with empty-textContent guard against picker false-positives.

---

## Architecture

```
┌────────────────────────────────────────────────────────────────────────────┐
│                            Auto Assign Bot                                 │
│                                                                            │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────────────────────┐  │
│  │ APScheduler  │ →  │ Playwright   │ →  │       ChangeGear ITSM        │  │
│  │ every 15 min │    │  Async API   │    │     (ASP.NET / IIS / AD)     │  │
│  └──────────────┘    └──────────────┘    └──────────────────────────────┘  │
│                              │                                             │
│              ┌───────────────┼──────────────┬──────────────┐               │
│              ↓               ↓              ↓              ↓               │
│       ┌────────────┐  ┌────────────┐ ┌───────────┐ ┌───────────────┐       │
│       │ Feedback   │  │ Claude AI  │ │ CMDB DB   │ │ History DB    │       │
│       │ (manual    │  │  Anthropic │ │ (SQLite)  │ │ (SQLite)      │       │
│       │  via       │  │  Sonnet 4.5│ │ asset     │ │ past tickets  │       │
│       │  teach.py) │  │  + caching │ │ owners    │ │ + corrections │       │
│       └────────────┘  └────────────┘ └───────────┘ └───────────────┘       │
│                              │                                             │
│                              ↓                                             │
│                       ┌─────────────┐                                      │
│                       │ Excel rules │  ← deterministic fallback            │
│                       │  (openpyxl) │                                      │
│                       └─────────────┘                                      │
└────────────────────────────────────────────────────────────────────────────┘
```

---

## Tech Stack

| Layer | Tech |
|---|---|
| Language | Python 3.11+ |
| Browser automation | Playwright async API (Chromium, headless-capable) |
| Popup capture | `BrowserContext.expect_page()` for CMDB scraping |
| Scheduling | APScheduler (`AsyncIOScheduler`) — default 15-min scan, 60-min correction sweep |
| Storage | SQLite 3 (two DBs: history + CMDB owners; feedback table in history DB) |
| AI | Anthropic Claude API (default `claude-sonnet-4-5`, configurable) |
| Prompt caching | `cache_control: {"type": "ephemeral"}` on stable system prompt blocks |
| Fuzzy match | Python `difflib.SequenceMatcher` (with trust weighting) |
| Config | Excel via openpyxl |
| HTTP auth | Playwright `http_credentials` (NTLM / Basic for AD) |
| Target system | ChangeGear ITSM (ASP.NET + AngularJS + DevExtreme + Telerik) |

---

## Decision Flow

Every new ticket runs through the same pipeline. Higher tiers short-circuit the lower ones.

```
1. Pre-compute Excel keyword → Requester's Item → CMDB Direct candidate
   (this becomes a fallback used by both AI and traditional paths)

2. Pull top-5 relevant feedback entries (SequenceMatcher vs ticket content)
   These become "Lessons" injected into the AI prompt.

──── AI mode (Claude API key set) ────

3. Claude receives ONE prompt containing:
   - Cacheable system: role + rules + valid Incident Types + valid CMDB CIs + JSON format
   - Volatile user message: feedback lessons + new ticket + top-N historical candidates,
     EACH candidate annotated with "CMDB comparison: ✓ matches / ✗ mismatches"

4. Claude returns owner / assigned_to / req_item / decision_source / confidence

   ├─ confidence  > 0.85 → dispatch immediately
   ├─ confidence == 0.85 → dispatch (CMDB validation skipped — Claude already saw CMDB)
   └─ confidence  < 0.85 → fall back to CMDB Direct candidate, else track-only

──── Traditional mode (no Claude key) ────

5. DB SequenceMatcher score
   ├─ CMDB Direct score (0.80) > DB score → use CMDB Direct
   ├─ DB score > 0.85 → dispatch immediately
   ├─ 0.75 ≤ DB score ≤ 0.85 → CMDB validation
   │     pass  → dispatch
   │     fail  → extract "Hi xxx" greeting from description → re-query DB (≥ 0.85) → dispatch
   ├─ 0.65 ≤ DB score < 0.75 → dispatch (low confidence but acceptable)
   └─ DB score < 0.65 → fall through to Excel keyword rules, then default
```

---

## Three-Signal Unified Decision (the Claude path explained)

Earlier versions of this bot used CMDB as a post-Claude validation override. That meant Claude couldn't *see* the CMDB ownership when deciding — leading to multi-stage fallback dances (Claude → CMDB validation → Hi-xxx name extraction → second-match…).

The current version puts everything in front of Claude in one call:

```
[For each historical candidate, the bot does cmdb_lookup(candidate.req_item) and inlines the result:]

[1] Ticket:IR-0093421 similarity:0.78 source:DB[manual-correction]
    Subject: VPN connection failure
    Requester: john@example.com
    Owner: Help Desk / Assigned To: Alice
    Incident Type: Service Request > Account Management
    Requester's Item: VPN_Access
    CMDB compare: CI=VPN_Access → Owner(team)=Infra / Owner(person)=Alice [✓ matches]

[2] Ticket:IR-0091234 similarity:0.71 source:DB[bot-dispatched]
    ...
    Requester's Item: Outlook_Mailbox
    CMDB compare: CI=Outlook_Mailbox → Owner(team)=Apps / Owner(person)=Bob [✗ mismatch]
```

Claude is then instructed:

- ✓ matches → safe to use the historical assigned_to
- ✗ mismatch → CMDB is usually more current (asset changes hands; history goes stale) → use CMDB owner
- Not in CMDB → use the most-similar historical candidate
- The ticket explicitly names a CMDB asset → CMDB owner takes precedence

The response includes a `decision_source` field (`feedback` / `cmdb` / `history` / `content` / `hybrid`) so logs show which signal actually drove the call.

---

## Operator Feedback Loop (`teach.py`)

The implicit correction-scan can detect *that* a human changed a bot dispatch, but not *why*. `teach.py` is the explicit channel.

### When to use it

- The bot dispatched a ticket to the wrong person and you fixed it manually
- You want the bot to learn the **principle** behind the correction, not just that one ticket

### Workflow

```bash
# Double-click Teach Bot.bat, or run from terminal:
python teach.py                  # interactive — asks for ticket id
python teach.py IR-0094069       # direct — go straight to the lesson form
python teach.py --list           # review all recorded lessons
python teach.py --delete <id>    # remove a wrong lesson
```

Sample session:

```
工單 ID: IR-0094200

Found ticket:
  ID:           IR-0094200
  Subject:      Re: Outlook group membership management
  Requester:    foo@example.com
  Owner:        Help Desk          ← what the bot picked
  Assigned To:  kyle.wu 吳鎧       ← wrong person

正確的 Assigned To: wen.hsieh
正確的 Owner: Help Desk

Why? (the reason becomes a lesson the bot will read on future similar tickets)
> Outlook group management is Exchange admin work, not generic Help Desk

✅ Recorded to feedback table.
```

### How the bot uses it

On the next Claude call, `db_get_relevant_feedback()` does a SequenceMatcher comparison between the new ticket content and every stored feedback entry. Top 5 entries with similarity ≥ 0.30 get inlined into the user message as a "## ⚠ 人工教學紀錄" section:

```
[Lesson 1] Past ticket IR-0094200 (similarity 0.82)
    Subject: Outlook group membership management
    Requester's Item: O365 - Exchange Online
    ✗ Wrong: Owner=Help Desk / Assigned=kyle.wu 吳鎧
    ✓ Correct: Owner=Help Desk / Assigned=wen.hsieh
    Rationale: Outlook group management is Exchange admin work, not generic Help Desk
```

The system prompt is configured to treat feedback as **rule 0** (higher priority than the three signals). When Claude applies a lesson, it returns `decision_source: "feedback"` for traceability.

---

## CMDB Validation (for borderline traditional-mode confidence)

In traditional mode (no Claude API key), when DB confidence falls in the 0.75–0.85 range, the bot cross-validates the dispatch choice against the CMDB asset's real owners:

```
Proposed Assigned To  ←→  CMDB.owner       (asset's individual owner)
Proposed Owner        ←→  CMDB.team_owner  (asset's team owner)
```

Match uses loose comparison (substring containment OR `SequenceMatcher` ratio ≥ 0.75). Both sides must match to pass.

If validation fails, the bot scans the ticket Description for greetings (`Hi xxx`, `Hello xxx`, `Dear xxx`) and does a second DB lookup by name, requiring score ≥ 0.85 to proceed.

> In AI mode this layer is unnecessary because Claude already sees CMDB info per candidate.

---

## Confidence Score & Help Desk Filter

Every dispatch decision carries an internal `_score` (range 0.0–1.0). Two policies use it:

### Help Desk owner filter

By default the bot only dispatches tickets whose Owner is "Help Desk", to avoid hijacking other teams' tickets. **Exception**: when `_score ≥ 0.75`, this filter is bypassed — the bot trusts the high-confidence match and dispatches regardless of Owner.

### Score sources and thresholds

| Source | `_score` | Triggers Help Desk bypass | Notes |
|---|---:|:---:|---|
| Claude AI (high confidence) | 0.85+ | ✅ | Confidence > 0.85 also skips CMDB validation |
| **CMDB Direct** | **0.80** | ✅ | Between the validation threshold and AI threshold |
| History DB high-confidence | ≥ 0.85 | ✅ | Skips CMDB validation |
| History DB mid-confidence | 0.75–0.85 | ✅ | Triggers CMDB validation |
| History DB low-confidence | 0.65–0.75 | ❌ | Dispatches but stays in Help-Desk-only mode |
| Excel keyword match | 0.00 | ❌ | |
| Default values | 0.00 | ❌ | |

---

## Weekday-Based SLA

Both Impact/Urgency/Priority and Due Date are derived from the day-of-week the ticket is dispatched:

| Dispatch day | Impact | Urgency | Priority | Due Date |
|---|---|---|---|---|
| Mon – Wed | 3 - Minor | 3 - Medium | 3 - Medium | +4 business days |
| Thu – Sun | 4 - Routine | 4 - Low | 4 - Low | +5 business days |

Business-day calculation skips weekends. Logic lives in `calc_business_due_date()`.

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
| `程式設定` (Settings) | AD account/password, ChangeGear URL, Claude API key (optional), scan interval, similarity threshold, Claude model, candidate count |
| `關鍵字派單規則` (Keyword Rules) | Keyword → Owner / Assigned To / Incident Type mapping |
| `Requester's Item 對應` | Keyword → Requester's Item mapping (used to derive CMDB candidates) |

The working copy is gitignored; only the `.example.xlsx` template is tracked.

### 3. (Optional) Customize TARGET_ASSIGNEES in `build_history_db.py`

If you want to broaden DB coverage beyond Help Desk owners, edit:

```python
TARGET_ASSIGNEES = ("user_a", "user_b", "user_c")  # ← lowercase AD account substrings
```

Used for the "Hi xxx" second-match learning.

### 4. Build the databases (first-time setup)

```bash
./Data\ base\ build.bat            # Windows — scrapes All Incidents into changegear_history.db
python build_history_db.py         # Cross-platform

./Build\ CMDB\ DB.bat              # Windows — scrapes All Managed Items into cmdb_owners.db
python build_cmdb_db.py            # Cross-platform
```

`build_history_db.py` honors a `MIN_TICKET_DATE = "2024-01-01"` cutoff by default — only scrapes tickets dated 2024 or later (the date is auto-detected from any cell in the row). Tickets with `Owner = Help Desk` or `Assigned To ∈ TARGET_ASSIGNEES` are also kept regardless of date.

Ctrl+C is safe — already-scraped rows are preserved.

### 5. Run the bot

```bash
./Auto\ mission\ start.bat
python changegear_auto_assign_v6.py
```

The scheduler scans for new tickets every N minutes (default 15) and runs an hourly correction-learning sweep.

### 6. Teach the bot when it makes a mistake

```bash
./Teach\ Bot.bat                   # Windows — interactive
python teach.py                    # Cross-platform interactive
python teach.py IR-0094069         # Direct
python teach.py --list             # List all lessons
python teach.py --delete <id>      # Remove a lesson
```

---

## Repository Layout

```
changegear-auto-assign/
├── changegear_auto_assign_v6.py        # Main dispatcher
├── build_history_db.py                 # Historical ticket scraper (2024+ by default)
├── build_cmdb_db.py                    # CMDB All-Managed-Items scraper
├── teach.py                            # Interactive feedback recorder
│
├── Auto mission start.bat              # Windows launcher: run dispatcher
├── Data base build.bat                 # Windows launcher: build history DB
├── Build CMDB DB.bat                   # Windows launcher: build CMDB DB
├── Teach Bot.bat                       # Windows launcher: teach.py
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
| `changegear_history.db` | Historical dispatch records + feedback table |
| `cmdb_owners.db` | CMDB asset owners |
| `auto_assign.log` | Dispatch decisions, decision_source, cache hit metrics, errors |
| `build_history.log` | History scraper log |
| `cmdb_build.log` | CMDB scraper log |
| `debug_*.png` | Screenshots of failed dispatch attempts |

---

## Configuration (Excel `程式設定` sheet)

| Parameter | Description | Default |
|---|---|---|
| `AD 帳號` | Windows AD account for ChangeGear login | — |
| `AD 密碼` | AD password | — |
| `系統網址` | ChangeGear base URL | (your CGWeb endpoint) |
| `掃描間隔（分鐘）` | Scan interval in minutes | 15 |
| `歷史比對相似度門檻` | DB similarity threshold (0–1) | 0.65 |
| `headless 模式` | True = run browser headlessly | False |
| `Claude API Key` | Anthropic API key (empty disables AI mode) | — |
| `Claude 模型` | Claude model ID | `claude-sonnet-4-5` |
| `Claude 候選工單數` | Historical candidates sent to Claude per call | 12 |
| `預設 Owner（無匹配時）` | Fallback owner when all layers miss | Help Desk |
| `預設 Assigned To（無匹配）` | Fallback assignee when all layers miss | IT-Helpdesk |

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

CREATE TABLE feedback (
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
    owner         TEXT,    -- maps to dispatch "Assigned To" (individual)
    co_owner      TEXT,
    tech_owner    TEXT,
    team_owner    TEXT,    -- maps to dispatch "Owner" (team)
    users_groups  TEXT,
    created_at    TIMESTAMP,
    updated_at    TIMESTAMP
);
```

---

## Self-Learning Loops

Two loops run alongside the main scan:

### Implicit (every hour) — silent UI correction watcher

- Scans tickets the bot dispatched (`bot_assigned = 1`) and tickets it tracked-but-didn't-dispatch (`bot_assigned = 2`).
- Detects when a human has edited the dispatch fields since the bot's action.
- Writes corrected values back to the DB and flags `corrected = 1`.
- Tickets with `corrected = 1` get the highest trust weight (× 1.20) in future similarity matches.

### Explicit (on-demand) — `teach.py`

- Operator decides which corrections matter and explains *why* in plain Chinese/English.
- Each lesson is stored with full context (summary, req_item, wrong vs correct dispatch, reason).
- Future Claude prompts pull the top-N most-relevant lessons and rank them above the three signals.

The two loops are complementary: the implicit one absorbs everyday corrections at scale; the explicit one captures the principles behind important corrections.

---

## Cost & Performance

For a typical 50-ticket/day workload running on `claude-sonnet-4-5` with 12 candidates and prompt caching:

| Metric | Value |
|---|---|
| API calls / month | ~1,100 (1 per dispatched ticket) |
| Input tokens / call (cached) | ~720 stable + ~3,000 volatile |
| Cache hit rate | ~95% on stable portion |
| Sonnet 4.5 effective cost | ~NT$ 350 / month |

If you switch to `claude-haiku-4-5`: roughly NT$ 120 / month at the same volume.

---

## FAQ

**Q: Why use Playwright popup capture for CMDB?**
The CMDB detail page (`SDItemEditPanel.aspx?boundtable=CIBase`) returns "Object reference not set" on direct navigation because server-side session state is only initialized when the page is opened via the module's `OnGetRowValues(oid)` JS function, which internally calls `window.open()`. Playwright's `BrowserContext.expect_page()` captures that popup cleanly.

**Q: Why two SQLite DBs instead of one?**
History DB grows continuously (one row per dispatched ticket) while CMDB DB is rebuilt periodically from the asset master list. Keeping them separate makes the rebuild cycle independent — you can re-scrape CMDB without losing dispatch history. Feedback shares the history DB since it's all "learning state".

**Q: Headless mode hangs on login.**
The ChangeGear server uses Windows Integrated Auth (NTLM). The bot passes `http_credentials` to Playwright's `BrowserContext`, which works in both headless and headed modes. If headless fails, verify your AD account isn't locked and check `auto_assign.log` for the actual error.

**Q: Why does Claude sometimes return `decision_source: "feedback"`?**
You (or another operator) recorded a relevant lesson via `teach.py`. Claude was instructed to treat feedback as the highest-priority signal. Check `python teach.py --list` to see which lessons are active.

**Q: I told the bot the wrong thing in `teach.py` — how do I remove it?**
`python teach.py --list` to find the id, then `python teach.py --delete <id>`.

**Q: Prompt caching seems to not be working.**
Check `auto_assign.log` for lines like `[Claude cache] read=720 write=0 uncached=2780 output=190`. `read > 0` means a hit; `write > 0` means first-time write or cache expired (5-minute TTL by default). If you see `read=0` on every call, something in the system prompt is changing per-request — likely the CMDB CI list or the Incident Type list grew. That's expected after `Build CMDB DB.bat` reruns; the cache rewarms within a few calls.

**Q: A ticket keeps reappearing in the "未派單" list even though logs show `✅ 完成`.**
Person Chooser autocomplete didn't actually match a real user. The fix is in place (AD account extraction + empty-textContent guard), but if you still see this:
- Manually verify the Assigned To field on one of those tickets in the ChangeGear UI
- Check whether the typed string matches an existing AD account format
- Watch for `⚠ ... dropdown 未出現` warnings in the log

**Q: How do I lower Claude API costs further?**
Three knobs:
- Reduce `Claude 候選工單數` (12 → 8 saves ~600 input tokens/call)
- Switch model to `claude-haiku-4-5` in Excel (3× cheaper, slightly lower accuracy)
- Lengthen the cache TTL by passing `{"type": "ephemeral", "ttl": "1h"}` in the system block (cache writes cost 2× instead of 1.25×, but stay alive 12× longer — break-even at 3+ calls per hour)

---

## Contributing

PRs welcome, especially:
- Adapters for other ITSM systems (ServiceNow, Jira Service Management, etc.)
- Alternative AI providers (OpenAI, local LLMs via Ollama)
- Tests / sample data fixtures
- Web UI for `teach.py` (currently CLI-only)

This was extracted from an internal tool — code quality is "works in production for one team" rather than "polished library". Suggestions for refactoring are appreciated.

---

## License

MIT — see [LICENSE](LICENSE).

> The author has no affiliation with SunView Software / ChangeGear. "ChangeGear" is a trademark of its respective owner. This project is an independent automation tool that interacts with ChangeGear via its standard web UI.
