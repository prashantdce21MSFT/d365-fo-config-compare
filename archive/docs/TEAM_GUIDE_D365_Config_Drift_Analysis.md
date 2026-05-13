# D365 Configuration Drift Analysis — Team Guide
## Using GitHub Copilot (Claude Code) + D365 MCP for Automated Form Comparison

**Author:** Prashant Verma, QA — Acme Asia SmartCore  
**Last Updated:** 2026-04-30  
**Project:** MY/SG Configuration Drift Analysis (UAT vs Env2 Golden)

---

## Table of Contents

1. [What This Does](#1-what-this-does)
2. [Prerequisites & Setup](#2-prerequisites--setup)
3. [Architecture Overview](#3-architecture-overview)
4. [End-to-End Workflow](#4-end-to-end-workflow)
5. [Step-by-Step: Running a Comparison](#5-step-by-step-running-a-comparison)
6. [How the Code Works (Under the Hood)](#6-how-the-code-works-under-the-hood)
7. [Known Gotchas & Workarounds (Critical)](#7-known-gotchas--workarounds-critical)
8. [Excel Output Format](#8-excel-output-format)
9. [Adding a New Form / New Project](#9-adding-a-new-form--new-project)
10. [Memory System — How AI Learns From Mistakes](#10-memory-system--how-ai-learns-from-mistakes)
11. [Troubleshooting](#11-troubleshooting)
12. [File Reference](#12-file-reference)

---

## 1. What This Does

This toolchain **automatically extracts and compares D365 F&O form configurations** between two environments (e.g., UAT vs Env2 Golden) across multiple legal entities (MY30, MY60, SG60).

**The problem it solves:** Manually opening 600+ forms across 2 environments x 3 legal entities, comparing every field, grid, tab, and dropdown — that's 3600+ manual checks. This tool does it programmatically.

**What it produces:** Excel files with side-by-side data from both environments, one row per record, filterable by Environment and Legal Entity. You visually scan for differences.

**How it talks to D365:** Via the **MCP (Model Context Protocol)** server embedded in D365 F&O. This is the same server that GitHub Copilot for Dynamics uses. Our code calls it directly via HTTP + OAuth2.

---

## 2. Prerequisites & Setup

### 2.1 Software Required

| Tool | Purpose | Install |
|------|---------|---------|
| **GitHub Copilot** (Claude Code / GHCP) | AI assistant that writes and runs the code | VS Code extension or CLI |
| **Python 3.10+** | Runtime for all scripts | python.org or winget |
| **Azure CLI** (`az`) | For ADO authentication (fetching work items) | `winget install Microsoft.AzureCLI` |
| **Git** | Version control (optional but recommended) | `winget install Git.Git` |

### 2.2 Python Packages

```bash
pip install openpyxl requests msal
```

### 2.3 Folder Structure

```
C:\D365DataValidator\                    # Shared library (reusable across projects)
    config.json                          # Environment configs (OAuth creds, MCP URLs)
    d365_mcp_client.py                   # MCP client class

C:\D365 Configuration Drift Analysis\   # Project-specific scripts
    form_control_extractor.py            # Main extraction engine
    compare_form_controls.py             # UAT vs Config comparison + Excel
    form_reader.py                       # MCP form tool wrappers
    run_batch_compare.py                 # Multi-form batch runner
    fetch_all_ado_items.py               # ADO work item extraction
    categorize_items.py                  # Map CDDs to D365 forms
    output\                              # All generated files go here
        Comparison\                      # Excel comparison files
        form_data\                       # JSON extraction backups
```

### 2.4 config.json Structure

**Location:** `C:\D365DataValidator\config.json`

```json
{
  "environments": {
    "ENV1": {
      "name": "Env1 UAT",
      "mcp_url": "https://example-uat.sandbox.operations.dynamics.com/mcp",
      "tenant_id": "00000000-0000-0000-0000-000000000000",
      "client_id": "00000000-0000-0000-0000-000000000001",
      "client_secret": "REDACTED"
    },
    "ENV4": {
      "name": "Env4 Config",
      "mcp_url": "https://example-config.sandbox.operations.dynamics.com/mcp",
      "tenant_id": "00000000-0000-0000-0000-000000000000",
      "client_id": "00000000-0000-0000-0000-000000000001",
      "client_secret": "REDACTED"
    }
  },
  "settings": {
    "max_records_per_entity": 50000,
    "request_timeout_seconds": 120,
    "retry_count": 3,
    "retry_delay_seconds": 5
  }
}
```

> **IMPORTANT:** The `tenant_id` must be the **Acme corporate tenant** (`00000000-0000-0000-0000-000000000000`), NOT your personal Azure tenant. The `client_id` and `client_secret` are for the app registration that has access to the D365 MCP endpoint.

### 2.5 Environment Reference

| Key | Name | URL | Purpose |
|-----|------|-----|---------|
| ENV1 | Env1 UAT | `example-uat.sandbox.operations.dynamics.com` | UAT environment (has real data) |
| ENV4 | Env4 Config | `example-config.sandbox.operations.dynamics.com` | Env2 Golden (reference baseline) |
| ENV2 | Env2 Golden EU | `example-goldconfig.sandbox.operations.eu.dynamics.com` | European Env2 Golden |
| ENV3 | Env3 UAT | `example-uat-eu.sandbox.operations.eu.dynamics.com` | Env3 UAT |

---

## 3. Architecture Overview

```
+------------------+     +-------------------+     +-------------------+
|  Azure DevOps    |     | D365 F&O UAT      |     | D365 F&O Config   |
|  (606 CDDs)      |     | (MCP Server)      |     | (MCP Server)      |
+--------+---------+     +--------+----------+     +--------+----------+
         |                         |                         |
         v                         v                         v
+--------+---------+     +--------+----------+     +--------+----------+
| fetch_all_ado    |     | D365McpClient     |     | D365McpClient     |
| _items.py        |     | (OAuth2 + MCP)    |     | (OAuth2 + MCP)    |
+--------+---------+     +--------+----------+     +--------+----------+
         |                         |                         |
         v                         v                         v
+--------+---------+     +--------+----------------------------+--------+
| categorize       |     |        FormControlExtractor                  |
| _items.py        |     |  - Opens form (form_open_menu_item)         |
+--------+---------+     |  - Parses FormState (grids, fields, tabs)   |
         |               |  - Paginates grids (25 rows/page)           |
         v               |  - Selects each row (detail fields)         |
+--------+---------+     |  - Opens sub-tabs (hidden fields)           |
| CDD_Form_Paths   |     |  - Runs form_find_controls sweep            |
| .xlsx            |     |  - OData enrichment (CustTable FastTabs)    |
+------------------+     +--------+-----------------------------------++
                                   |                         |
                                   v                         v
                         +--------+----------+     +--------+----------+
                         | UAT Result (JSON) |     | Config Result      |
                         +--------+----------+     +--------+----------+
                                   |                         |
                                   +----------+--------------+
                                              |
                                              v
                                   +----------+----------+
                                   | compare_form        |
                                   | _controls.py        |
                                   | (to_excel)          |
                                   +----------+----------+
                                              |
                                              v
                                   +----------+----------+
                                   | Compare_FormName    |
                                   | _YYYYMMDD.xlsx      |
                                   +---------------------+
```

### Key Components

| Component | File | What It Does |
|-----------|------|-------------|
| **MCP Client** | `d365_mcp_client.py` | OAuth2 authentication + JSON-RPC calls to D365 MCP server |
| **Form Reader** | `form_reader.py` | Wrappers for MCP form tools (open, close, navigate, read) |
| **Extractor** | `form_control_extractor.py` | Orchestrates full form extraction (grids, pagination, tabs, sweeps) |
| **Comparison** | `compare_form_controls.py` | Runs extractor on 2 environments, generates Excel |
| **ADO Fetcher** | `fetch_all_ado_items.py` | Pulls CDD work items from ADO |
| **Categorizer** | `categorize_items.py` | Maps CDD titles to D365 menu item names |

### MCP Tools Used

The D365 MCP server exposes these tools (called via `client.call_tool()`):

| MCP Tool | Purpose |
|----------|---------|
| `form_find_menu_item` | Search for a form by name |
| `form_open_menu_item` | Open a form (returns FormState) |
| `form_close_all_forms` | Close all open forms |
| `form_select_grid_row` | Select a specific grid row (returns updated FormState) |
| `form_click_control` | Click a control (used for `LoadNextPage` pagination) |
| `form_open_or_close_tab` | Open/close a tab section |
| `form_find_controls` | Search for controls by name substring |
| `data_find_entity_type` | Search OData entities (fuzzy match) |
| `data_get_entity_metadata` | Get entity schema (fields, types, keys) |
| `data_find_entities` | Query OData (like a REST GET with $filter, $select) |

---

## 4. End-to-End Workflow

This is the full flow from "I have 606 CDDs in ADO" to "I have comparison Excel files":

### Phase 1: ADO Extraction (One-time Setup)

**Goal:** Get all CDD work items from ADO, understand what forms to validate.

```
Step 1: Run fetch_all_ado_items.py
        - Reads ADO query (7a10e4a5-...) via REST API
        - Downloads 606 CDD items with titles, descriptions, screenshots
        - Output: ado_items.json, output/screenshots/

Step 2: Run generate_cdd_excel.py
        - Parses navigation paths from CDD titles
        - Maps [MY]/[SG] tags to company codes
        - Output: CDD_Form_Paths_YYYYMMDD.xlsx

Step 3: Run categorize_items.py
        - For each CDD, searches D365 for matching form (MCP + OData)
        - Scores matches, picks best menu item name
        - Output: MY_SG_606_Categorization.xlsx
        - This gives you: ADO ID -> Menu Item Name + Type
```

### Phase 2: Form-by-Form Comparison (Ongoing Work)

**Goal:** For each form, extract from UAT and Config, generate Excel.

```
Step 4: Edit compare_form_controls.py (or ask GHCP to do it)
        - Set mi_name = "InventLocations"  (or whatever form)
        - Set legal_entities = ["MY30", "MY60", "SG60"]
        - Run it

Step 5: Script extracts from ENV1 (UAT), then ENV4 (Config)
        - Opens form, paginates grids, selects each row
        - Opens all sub-tabs, captures detail fields
        - Runs form_find_controls sweep

Step 6: Generates comparison Excel
        - output/Comparison/Compare_FormName_YYYYMMDD_HHMMSS.xlsx
```

### Phase 3: Batch Comparison (Optional)

```
Step 7: Use run_batch_compare.py for multiple forms at once
        - Define FORMS list with mi_name, mi_type, ado_id
        - Runs Phase 2 for each form sequentially
```

---

## 5. Step-by-Step: Running a Comparison

### Option A: Ask GHCP to Do It (Recommended)

Open your GHCP session in the project directory and say:

> "Extract and compare CustTable form between UAT (ENV1) and Config (ENV4) for legal entities MY30, MY60, SG60. Generate comparison Excel."

GHCP will:
1. Read the existing `compare_form_controls.py`
2. Modify the `mi_name`, `mi_type`, `legal_entities` variables
3. Run the script
4. Report the output Excel path

### Option B: Run Manually

1. Edit `compare_form_controls.py`:
```python
# At the bottom of the file, in the __main__ block:
legal_entities = ["MY30", "MY60", "SG60"]
mi_name = "InventLocations"      # <-- Change this
mi_type = "Display"               # <-- Display, Action, or Output
```

2. Run:
```bash
cd "C:\D365 Configuration Drift Analysis"
set PYTHONUNBUFFERED=1
python compare_form_controls.py
```

3. Output will be at: `output/Comparison/Compare_InventLocations_YYYYMMDD_HHMMSS.xlsx`

### How Long Does It Take?

| Scenario | Time |
|----------|------|
| Form with 0 grid rows (Config typically) | ~2-3 minutes (sweep only) |
| Form with 25 rows, 1 tab | ~3-5 minutes |
| Form with 50 rows, 5 tabs | ~5-10 minutes |
| Form with 300 rows, 1 tab (CustTable MY60) | ~15-20 minutes |
| Full CustTable (3 LEs, UAT+Config) | ~25-35 minutes |

Most time is spent on per-row detail extraction (selecting each row, opening tabs).

---

## 6. How the Code Works (Under the Hood)

### 6.1 MCP Communication

```
Your Script
    |
    v
D365McpClient._mcp_request()
    |  POST https://example-uat.../mcp
    |  Headers: Authorization: Bearer {oauth_token}
    |           Mcp-Session-Id: {session_id}
    |  Body: {"jsonrpc":"2.0","method":"tools/call",
    |         "params":{"name":"form_open_menu_item",
    |                   "arguments":{"menuItemName":"CustTable",...}}}
    v
D365 MCP Server
    |  Opens the form server-side, returns FormState JSON
    v
FormState JSON (huge nested structure):
{
  "FormState": {
    "Name": "CustTable",
    "Caption": "Customers",
    "Form": {
      "Tab": {
        "TabPageGrid": {
          "Children": {
            "Grid": {
              "Grid": {
                "Columns": [...],
                "Rows": [...],          // max 25 per page
                "Pagination": {"HasNextPage": "True"}
              }
            }
          }
        }
      },
      "Input": { ... },
      "Combobox": { ... }
    }
  }
}
```

### 6.2 Extraction Pipeline (Per Legal Entity)

```
1. CONNECT
   - Acquire OAuth2 token (client credentials grant)
   - Initialize MCP session (protocol version 2024-11-05)
   - Retry up to 3x with 10s/20s backoff

2. OPEN FORM
   - form_open_menu_item(menuItemName, menuItemType, companyId)
   - Parse initial FormState -> grids, fields, tabs

3. PAGINATE GRIDS
   - For each grid with 25+ rows:
     - form_click_control(controlName=gridName, actionId="LoadNextPage")
     - Grid may be NESTED inside Tabs after pagination (recursive search)
     - Repeat until HasNextPage = "False"

4. EXTRACT ROW DETAILS
   - For each grid row (in groups of 25):
     - form_select_grid_row(gridName, rowNumber)
     - Parse returned FormState for detail fields (General tab)
     - For each sub-tab: form_open_or_close_tab(tabName, tabAction="Open")
     - Parse sub-tab FormState for additional fields
   - After every 25 rows: close and reopen form, paginate back
     (MCP state corrupts after ~25 consecutive row selections)

5. ODATA ENRICHMENT (CustTable only)
   - form_find_controls returns STALE values for FastTab fields
   - Query OData CustomersV3 entity with $filter by CustomerAccount+dataAreaId
   - Merge 32 CBA/CHB fields (Consignment warehouse, Sales office, etc.)
   - Match grid rows by CustomerAccount column value

6. FORM_FIND_CONTROLS SWEEP
   - Search with 100+ terms (a-z, digits, common prefixes)
   - Captures controls NOT in the initial FormState
   - Deduplicates against already-captured fields

7. RETURN
   - Nested dict: grids (with rows + row_details), fields, tabs, summary
```

### 6.3 Why Close/Reopen Every 25 Rows?

The MCP server maintains form state server-side. After ~20-25 consecutive `form_select_grid_row` calls, the server starts returning empty/error responses (`Expecting value: line 1 column 1`). The workaround:

1. Close the form (`form_close_all_forms`)
2. Reopen it (`form_open_menu_item`)
3. Paginate back to the correct page (repeat `LoadNextPage` N times)
4. Continue from where we left off

This is automatic — the extractor handles it transparently.

---

## 7. Known Gotchas & Workarounds (CRITICAL)

These were discovered through painful debugging. **Read this before modifying any code.**

### 7.1 MCP Parameter Names (Silent Failures)

| Tool | Wrong | Correct |
|------|-------|---------|
| `form_find_controls` | `searchTerm` | `controlSearchTerm` |
| `form_select_grid_row` | `rowNumber: 1` (int) | `rowNumber: "1"` (string) |
| `form_open_or_close_tab` | no tabAction | `tabAction: "Open"` (required!) |
| `data_find_entities` | `entityName` | `odataPath` |
| `data_get_entity_metadata` | `entityName` | `entitySetName` |
| `data_find_entity_type` | `searchTerm` | `tableSearchFilter` |

**These cause SILENT failures** — the MCP server returns valid JSON but with empty/wrong data. No error message.

### 7.2 Grid Pagination

- D365 MCP returns **max 25 rows per page**
- After `LoadNextPage`, grid data stays **nested inside Tab/TabPage** hierarchy
- Must search recursively (not just `Form > Grid > GridName`)
- Check `Pagination.HasNextPage` — but even if "False" with exactly 25 rows, try one more page
- `HasNextPage` is a **string** (`"True"/"False"`), not boolean

### 7.3 form_open_or_close_tab Requires tabAction

```python
# WRONG — silently fails, returns error in Result but no FormState
client.call_tool("form_open_or_close_tab", {"tabName": "General"})

# CORRECT
client.call_tool("form_open_or_close_tab", {"tabName": "General", "tabAction": "Open"})
```

Without `tabAction`, the extractor captures only ~9-12 fields per row instead of ~56-62.

### 7.4 form_find_controls Returns Duplicates

`form_find_controls` returns **3 copies** of each control:
- 1st copy: updates per row but often shows EMPTY for non-first rows
- 2nd/3rd copies: retain STALE values from initial form load

**Impact:** Per-row `form_find_controls` gives wrong values. Don't rely on it for per-row FastTab fields.

**Workaround:** Use OData for per-record field values (see CustTable approach).

### 7.5 OData Entity Matching (Fuzzy Search Danger)

`data_find_entity_type` does **fuzzy search** — it can return wrong entities.

**Always verify before using OData:**
1. **Table match** — control name `CustTable_CHBConsignmentWarehouse` tells you table=`CustTable`. Verify OData entity maps to same table.
2. **Field match** — check `get_entity_metadata` has the exact field name
3. **Value cross-check** — query one known record, compare against what the form shows

### 7.6 OAuth Timeouts Between Legal Entities

When switching from MY30 to MY60, the new MCP client connection can timeout. The code retries 3x with 10s/20s backoff automatically.

### 7.7 Windows cp1252 Encoding

Box-drawing characters (----, ====) crash on Windows console. All scripts use:
```python
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
```

### 7.8 D365 Control Name Pattern

Controls follow `TableName_FieldName` pattern:
- `CustTable_CHBConsignmentWarehouse` -> Table: `CustTable`, Field: `CHBConsignmentWarehouse`
- `BOMTable_Name` -> Table: `BOMTable`, Field: `Name`
- Use this to identify which table a form reads from

---

## 8. Excel Output Format

### Summary Sheet

| Field | Value |
|-------|-------|
| Form | CustTable |
| MI Type | Display |
| UAT | ENV1 (Env1 UAT) |
| Config | ENV4 (Env4 Config) |
| Legal Entities | MY30, MY60, SG60 |
| MY30 UAT rows | 47 |
| MY30 Config rows | 0 |
| MY60 UAT rows | 296 |
| ... | ... |

### Grid Sheet (one per grid)

```
| # | Customer account | Name         | Group | ... | [D] (Acme Asia) > Consignment warehouse | [D] (General) > Site | Env    | LE   |
|---|------------------|--------------|-------|-----|----------------------------------------------|----------------------|--------|------|
|   | MY30             |              |       |     | UAT: 47 rows                                 | Config: 0 rows       |        |      |  <- separator
| 1 | 210205           | ABC Trading  | 30    | ... | CON                                          | 61                   | UAT    | MY30 |
| 2 | 210300           | XYZ Corp     | 30    | ... | FG                                           | 61                   | UAT    | MY30 |
| ...                                                                                                                                      |
|   | MY60             |              |       |     | UAT: 296 rows                                | Config: 0 rows       |        |      |  <- separator
| 48| 1000             | Test Cust    | 10    | ... | CON                                          | 62                   | UAT    | MY60 |
```

**Column conventions:**
- Grid columns appear first (matching D365 Standard view)
- Detail fields (from row selection / OData) prefixed with `[D]` and tab name: `[D] (Acme Asia) > Consignment warehouse`
- `Env` column = "UAT" or "Config"
- `LE` column = company code (MY30, MY60, SG60)
- Separator rows with LE + row counts for quick scanning
- Auto-filter enabled

---

## 9. Adding a New Form / New Project

### 9.1 Adding a New Form to Compare

1. Find the Menu Item name (look in `batch_extract_mi_cache.json` or ask GHCP)
2. Edit `compare_form_controls.py`:
```python
mi_name = "YourFormName"    # e.g., "InventLocations", "WHSInventStatus"
mi_type = "Display"          # Display, Action, or Output
legal_entities = ["MY30", "MY60", "SG60"]
```
3. Run the script

### 9.2 Adding a New Environment

1. Get the MCP URL from the D365 environment (format: `https://{env-name}.sandbox.operations.dynamics.com/mcp`)
2. Get OAuth credentials (app registration with D365 access)
3. Add to `config.json`:
```json
"ENV5": {
  "name": "New Environment",
  "mcp_url": "https://new-env.sandbox.operations.dynamics.com/mcp",
  "tenant_id": "00000000-0000-0000-0000-000000000000",
  "client_id": "00000000-0000-0000-0000-000000000001",
  "client_secret": "your-secret"
}
```
4. Update `compare_form_controls.py` to use the new env key

### 9.3 Starting a New Project (e.g., Region3, HUB)

1. Create a new project directory (e.g., `C:\D365Region3Validator\`)
2. Copy from REGION project:
   - `form_control_extractor.py`
   - `compare_form_controls.py`
   - `form_reader.py`
3. The shared library (`C:\D365DataValidator\`) is reusable — just point to it
4. Update `config.json` with the new environment URLs
5. Update legal entities in your comparison scripts

---

## 10. Memory System — How AI Learns From Mistakes

This is the key differentiator. When you use GHCP for this work, it **remembers what went wrong and what worked**.

### How It Works

GHCP stores "memories" in `.claude/projects/.../memory/` as markdown files. These are loaded into every new conversation, so the AI doesn't repeat mistakes.

### What's Stored

| Memory | What It Prevents |
|--------|-----------------|
| `feedback_mcp_form_gotchas.md` | Using wrong parameter names, missing `tabAction`, etc. |
| `feedback_form_tools_pagination.md` | Forgetting to paginate (truncating at 25 rows) |
| `feedback_odata_entity_matching.md` | Picking wrong OData entity from fuzzy search |
| `feedback_grid_excel_layout.md` | Generating unreadable flat key-value Excel |
| `feedback_tab_action_required.md` | Missing sub-tab fields (silent failure) |
| `feedback_form_table_extraction.md` | Wrong table name extraction approach |

### How to Build This for Your Session

When you start using GHCP on a new machine/account:

1. Copy the memory folder: `.claude/projects/c--Acme-SQL-Analaysis-for-DM/memory/` to your machine's equivalent path
2. Or: tell GHCP about each gotcha manually — it will save them as memories
3. Or: point GHCP to `_local_replica/WORKAROUNDS_AND_LESSONS.md` and say "remember all of this"

### Adding New Memories

When you discover a new bug/gotcha, tell GHCP:

> "Remember: form_find_controls returns 3 duplicate copies of each control. The first copy updates per row but is often empty. Don't use it for per-row field extraction."

GHCP will save this and apply it in all future conversations.

---

## 11. Troubleshooting

### Script hangs with no output
- Python buffers stdout. Run with `PYTHONUNBUFFERED=1`:
  ```bash
  set PYTHONUNBUFFERED=1
  python compare_form_controls.py
  ```

### "Expecting value: line 1 column 1" errors
- MCP server returned empty response. Usually happens after ~20-25 consecutive row selections.
- The code auto-recovers (mid-page recovery: close + reopen + paginate back).
- If it fails repeatedly, the environment may be under load — try again later.

### OAuth token error / 401 Unauthorized
- Check `config.json` credentials
- Ensure `tenant_id` is the Acme corporate tenant (`00000000-0000-0000-0000-000000000000`)
- App registration may need re-consent or secret rotation

### Form shows 0 rows in Config
- This is normal. Env2 Golden environments often have no transactional/master data.
- The comparison is still valuable for form-level fields (find_controls sweep).

### Wrong values in Excel
- Check which data source was used (form_find_controls vs OData vs FormState)
- For CustTable: OData `CustomersV3` is the correct source for Acme Asia FastTab fields
- For other forms: verify by manually opening the form in D365 and comparing one row

### "Menu item not found"
- Check the MI cache: `output/batch_extract_mi_cache.json`
- Try different search terms in D365 (form names != menu item names)
- Some forms are Actions, not Display — try `mi_type = "Action"`

---

## 12. File Reference

### Core Files (You'll Use These)

| File | Purpose | When to Modify |
|------|---------|---------------|
| `compare_form_controls.py` | Main comparison script | Change `mi_name`, `legal_entities` per form |
| `run_batch_compare.py` | Batch multiple forms | Add forms to `FORMS` list |
| `C:\D365DataValidator\config.json` | Environment credentials | New environments or secret rotation |

### Library Files (Don't Modify Unless Fixing Bugs)

| File | Purpose |
|------|---------|
| `form_control_extractor.py` | Extraction engine (39KB) — handles all MCP interactions |
| `form_reader.py` | MCP form tool wrappers (34KB) |
| `C:\D365DataValidator\d365_mcp_client.py` | OAuth + MCP transport (5.5KB) |

### ADO Integration Files (One-Time Use)

| File | Purpose |
|------|---------|
| `fetch_all_ado_items.py` | Pull CDDs from ADO query |
| `generate_cdd_excel.py` | Parse CDD titles into form paths |
| `categorize_items.py` | Map CDDs to D365 menu items |

### Reference

| File | Purpose |
|------|---------|
| `_local_replica/README.md` | Architecture overview |
| `_local_replica/WORKAROUNDS_AND_LESSONS.md` | 15 documented gotchas |
| `output/batch_extract_mi_cache.json` | Navigation path -> menu item name cache (46KB) |

---

## Quick Start Checklist

- [ ] Python 3.10+ installed with `openpyxl`, `requests`, `msal`
- [ ] `C:\D365DataValidator\config.json` has correct credentials
- [ ] `C:\D365 Configuration Drift Analysis\` folder exists with all scripts
- [ ] Test connectivity: run a small form (e.g., `Interest` with 1 LE) first
- [ ] Copy memory files to your `.claude/projects/` path (or let GHCP learn from `_local_replica/WORKAROUNDS_AND_LESSONS.md`)
- [ ] Run `compare_form_controls.py` with your target form

---

*For questions or issues, check the `_local_replica/WORKAROUNDS_AND_LESSONS.md` file first — most common problems are documented there.*
