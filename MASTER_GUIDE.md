# D365 FO Config Compare — Master Guide

> **The complete story.** Business problem, full project history, architecture, both backends (MCP & Playwright) in deep detail, every gotcha we hit, the original Python pipeline, and how it all became a VS Code extension.

This is **one document** consolidating everything from the D365 Configuration Drift Analysis project (Apr 28 – May 13, 2026). For shorter task-specific docs see [README.md](README.md), [SKILL.md](SKILL.md), and the [walkthrough](walkthrough/).

---

## Table of Contents

1. [Business Problem](#1-business-problem)
2. [Project Timeline](#2-project-timeline)
3. [Architecture](#3-architecture)
4. [Backends: MCP vs Playwright — Complete Comparison](#4-backends-mcp-vs-playwright--complete-comparison)
5. [D365 ERP MCP Backend — Deep Dive](#5-d365-erp-mcp-backend--deep-dive)
6. [Playwright (Chrome CDP) Backend — Deep Dive](#6-playwright-chrome-cdp-backend--deep-dive)
7. [Form-Path Validator](#7-form-path-validator)
8. [Output Format](#8-output-format)
9. [Complete Gotcha Catalogue (20+ entries)](#9-complete-gotcha-catalogue-20-entries)
10. [Authentication](#10-authentication)
11. [Original Python Pipeline (Pre-Extension Era)](#11-original-python-pipeline-pre-extension-era)
12. [VS Code Extension Architecture](#12-vs-code-extension-architecture)
13. [Governance & Usage Model](#13-governance--usage-model)
14. [Operational Runbook](#14-operational-runbook)
15. [Roadmap](#15-roadmap)
16. [File Reference (Old + New)](#16-file-reference-old--new)

---

## 1. Business Problem

Acme's REGION D365 F&O program — like most enterprise F&O programs — repeatedly hit the same class of late-cycle bug:

> *"This worked in DEV. In UAT it fails. Why?"*

The cause was almost always **configuration drift**: some form, parameter or master-data record differed between the **Env2 Golden** environment and **UAT / SIT / PROD-like** environments. These mismatches:

- caused hard-to-diagnose functional bugs late in the cycle,
- consumed multiple consultant-days per incident,
- and were invisible to backend diff tools — the difference was at the **form / parameter** level, not the entity level.

Manually opening **606 forms × 2 environments × 3 legal entities (MY30, MY60, SG60)** = **3,636 manual comparisons**. That is the work this project automates.

### Why not just compare Data Entities?

- Data entities scale to thousands of rows efficiently.
- **But** they produce hundreds of technical columns that functional consultants cannot map back to a form.
- The whole value proposition is **field-to-UI mapping**. Hence form-based extraction is non-optional.

### Audience

- **Functional consultants** — read the diff, decide the correct value, file the ADO ticket.
- **QA leads** — schedule extractions, gate UAT entry on a clean diff report.
- The output **must be form-shaped** (one sheet per tab/grid, field labels not internal names).

---

## 2. Project Timeline

| Phase | Date | What happened |
|---|---|---|
| **Phase 0 — ADO ingestion** | Apr 28 | Pulled 606 CDD work items from ADO query, downloaded screenshots, categorised each item to a D365 menu item. Files: `fetch_all_ado_items.py`, `categorize_items.py`, `MY_SG_606_Categorization.xlsx`. |
| **Phase 1 — Lessons crystallise** | Apr 29 | First wave of MCP gotchas discovered (parameter names, pagination nesting, tab actions). [`WORKAROUNDS_AND_LESSONS.md`](archive/docs/WORKAROUNDS_AND_LESSONS.md) authored. Debug scripts proliferate. |
| **Phase 2 — Team guide** | Apr 30 | [`TEAM_GUIDE_D365_Config_Drift_Analysis.md`](archive/docs/TEAM_GUIDE_D365_Config_Drift_Analysis.md) authored (.md + .docx + .pdf). **`form_control_extractor.py` (53.8 KB)** matures — this is the engine later ported into the VS Code extension. |
| **Phase 3 — Form/table mapping + mass extraction** | May 1 – May 4 | `build_form_mapping.py`, `build_tables_mapping.py`, `deep_extract_parameters.py`, `batch_extract_asset_fa.py`. 30+ Excel outputs produced for Asset & FA module. |
| **Phase 4 — MCP vs Playwright study** | May 11 | [`D365_ERP_MCP_vs_Playwright_Comparison.docx`](archive/docs/D365_ERP_MCP_vs_Playwright_Comparison.docx) — formal comparison report. Key finding: neither tool is universally better; use both selectively. |
| **Phase 5 — Extension inception** | May 11 (evening) | First VS Code extension scaffold. Publisher = "Prashant Verma AIBS". |
| **Phase 6 — Polish & ship** | May 12 – May 13 | Multiple chat sessions: description, error fixes, walkthrough scaffolding ("dataverse-onboarding-style"). Spawned sibling tool [C:\Bugs daily report] for ADO 6 AM SGT email automation. |
| **Phase 7 — Productisation** | May 13 | Validator hardening → multi-select → deep MCP port (5-stage pipeline) → rename to "D365 FO Config Compare" → new icon → Beta marking → public GitHub repo → SKILL.md → this master guide. |

---

## 3. Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│ VS Code Extension  (TypeScript host + Python extractor)             │
│                                                                     │
│   ┌──────────────────────────────────────────┐                      │
│   │ src/                                     │                      │
│   │   extension.ts        — commands         │                      │
│   │   extractor.ts        — multi-select     │                      │
│   │   onboarding.ts       — env + az login   │                      │
│   │   validator.ts        — path discovery   │                      │
│   │   welcome.ts          — walkthrough      │                      │
│   └──────────────────────────────────────────┘                      │
│                       │ spawns                                      │
│                       v                                             │
│   ┌──────────────────────────────────────────┐                      │
│   │ python/extract.py    (~800 lines)        │                      │
│   │   ├── extract_mcp          ◄── 5-stage   │                      │
│   │   │     ├── _parse_form    recursive     │                      │
│   │   │     ├── _expand_fasttabs             │                      │
│   │   │     ├── _paginate_grid               │                      │
│   │   │     └── _find_controls_sweep         │                      │
│   │   ├── extract_playwright   ◄── CDP       │                      │
│   │   │     ├── _JS_TABS                     │                      │
│   │   │     ├── _scroll_collect_grids        │                      │
│   │   │     └── _ACTION_BUTTONS (allowlist)  │                      │
│   │   ├── validate_form_paths                │                      │
│   │   └── write_excel + write_diff           │                      │
│   └──────────────────────────────────────────┘                      │
│              │                                │                     │
│              │ HTTPS + Bearer                 │ CDP :9222           │
│              v                                v                     │
│   ┌────────────────────────┐    ┌──────────────────────────────┐    │
│   │ D365 F&O /mcp endpoint │    │ User's signed-in Chrome      │    │
│   │ (JSON-RPC)             │    │ (logged into D365 F&O)       │    │
│   └────────────────────────┘    └──────────────────────────────┘    │
└─────────────────────────────────────────────────────────────────────┘
                       │
                       v
                ┌──────────────┐
                │ Excel output │ <form>_<le>_<env>_<ts>.xlsx
                │  + DIFF      │ <form>_<le>_DIFF_<ts>.xlsx
                └──────────────┘
```

### Settings persisted in `.vscode/settings.json`

```jsonc
{
  "d365FormExtractor.environments": [
    { "label": "Env1 UAT",       "baseUrl": "https://example-uat...",    "mcpAvailable": true },
    { "label": "Env4 Config",  "baseUrl": "https://example-config...", "mcpAvailable": true }
  ],
  "d365FormExtractor.formPaths": [
    { "path": "Accounts payable > Setup > Vendor groups", "menuItem": "VendGroup", "validated": true }
  ],
  "d365FormExtractor.defaultBackend": "ask"
}
```

No secrets. Auth happens through `az` on each run.

---

## 4. Backends: MCP vs Playwright — Complete Comparison

This section consolidates the May 11 study with what we learned during productisation.

### Strengths & weaknesses

| Dimension | **D365 ERP MCP** | **Playwright (Chrome CDP)** |
|---|---|---|
| **Output quality** | Structured, form-aware, tab-mapped | Same shape, sourced from DOM |
| **Speed (simple form)** | Fast — single JSON-RPC | Comparable |
| **Speed (complex form)** | Slower — many tool calls + sweep | Faster |
| **Headless?** | Yes | No (needs Chrome window) |
| **Auth model** | Azure CLI bearer (per env) | Existing Chrome session |
| **Setup overhead** | None (after `az login`) | Must launch Chrome with `--remote-debugging-port=9222` |
| **Misses segmented controls** | Yes (financial dimensions blank) | No — DOM shows rendered text |
| **Grid pagination** | 25-row default page, must drive | Virtual-scroll captures everything |
| **Response size limit** | ~160 KB per response | None |
| **Coverage on complex forms** | Lower in our trials (5 diffs vs 8) | Higher |
| **Risk surface** | Read-only by protocol design | UI-level — can trigger actions if not guarded |
| **Required prompts** | None | Strict allowlist forbidding Create/Edit/Delete |

### Verdict

> **Neither is universally superior. Let the user pick per form.**

The extension surfaces this choice at run-time (`D365 FO Config Compare: Extract Form` → backend picker). The recommendation is:

- **Default to MCP.** It's headless, deterministic, structured, and the deep extraction port now matches the original Python extractor's coverage.
- **Switch to Playwright** when MCP returns empty tabs, when the form has financial-dimension segmented controls, when you want to *see* the extraction, or when MCP isn't enabled on that tenant.

---

## 5. D365 ERP MCP Backend — Deep Dive

### 5.1 Transport

JSON-RPC 2.0 over HTTPS to `<env>/mcp`. Protocol version `2024-11-05`. Each session is initialised, then tool calls are made with `tools/call`. Session ID is carried in the `Mcp-Session-Id` header.

### 5.2 Tools used

| Tool | Required args | Purpose |
|---|---|---|
| `form_find_menu_item` | `searchTerm` | Locate a form by name |
| `form_open_menu_item` | `menuItemName`, `menuItemType`, `companyId` | Open a form — returns `FormState` JSON |
| `form_close_all_forms` | — | Cleanup |
| `form_open_or_close_tab` | `tabName`, **`tabAction`** ("Open" / "Close") | Force-open / collapse a tab |
| `form_click_control` | `controlName`, optional `actionId` | Activate a control. With `actionId="LoadNextPage"` triggers grid pagination |
| `form_select_grid_row` | `gridName`, **`rowNumber` (STRING)** | Select a row → returns FormState with row-detail fields |
| `form_find_controls` | `controlSearchTerm` | Search controls by substring |
| `data_find_entity_type` | `tableSearchFilter` | Search OData entities |
| `data_get_entity_metadata` | `entitySetName` | Entity schema (fields, types) |
| `data_find_entities` | `odataPath` | OData query (`$filter`, `$select`) |

**Each one of those parameter names** in the second column is a foot-gun. Wrong name → silent failure (valid JSON, empty/wrong data, no error). See § 9.

### 5.3 The 5-stage extraction pipeline (the heart of `extract_mcp`)

```
┌──────────────────────────────────────────────────────────────┐
│ STAGE 1 — Open form + initial recursive parse                │
│   form_open_menu_item(mi, type, le)                          │
│   _parse_form(FormState.Form, depth=0)                       │
│     ├── walks Tab → TabPage → Group → FastTab →              │
│     │   ReferenceGroup → Children recursively                │
│     ├── _parse_input / _parse_combobox / _parse_checkbox     │
│     ├── merges nested fields into top-level flat dict        │
│     └── preserves per-tab structure for sheet naming         │
├──────────────────────────────────────────────────────────────┤
│ STAGE 2 — Expand every collapsed FastTab                     │
│   _expand_fasttabs(client, fasttabs)                         │
│     └── form_click_control(name) per FastTab                 │
├──────────────────────────────────────────────────────────────┤
│ STAGE 3 — Force-open every Tab                               │
│   for each Tab:                                              │
│     try   form_open_or_close_tab(tabName, "Open")            │
│     catch form_click_control(tabName)  # fallback            │
│     re-parse active payload                                  │
│     attribute newly-found fields/grids to that tab           │
│   subtract tab-attributed fields from "top" (dedup)          │
├──────────────────────────────────────────────────────────────┤
│ STAGE 4 — Paginate every grid                                │
│   for each grid (top + per-tab):                             │
│     while HasNextPage == "True" OR rows >= 25:               │
│       form_click_control(grid, actionId="LoadNextPage")      │
│       _find_grid_recursive(payload, gridName)  # nested!     │
│       dedup rows by tuple-key of column values               │
│     (up to 200 pages safety cap)                             │
├──────────────────────────────────────────────────────────────┤
│ STAGE 5 — Hidden-control sweep                               │
│   _find_controls_sweep iterates ~146 terms:                  │
│     a-z, 0-9, plus ~100 domain terms                         │
│     (Code, Name, Date, Account, Vend, Cust, Item, Posting,   │
│      Default, Tax, Currency, Group, Setup, Parameters, …)    │
│   form_find_controls(controlSearchTerm=term)                 │
│   add any non-grid controls missing from stages 1-3          │
│   skip SystemDefined* prefix                                 │
│   Output: "Sweep added N new field(s)"                       │
└──────────────────────────────────────────────────────────────┘
```

### 5.4 The 25-row close/reopen recovery (from original Python, **not yet in extension**)

The original `form_control_extractor.py` discovered that after ~20-25 consecutive `form_select_grid_row` calls the MCP server returns `Expecting value: line 1 column 1` errors — server-side state corrupts. Workaround:

1. `form_close_all_forms`
2. `form_open_menu_item` again
3. Paginate back to the correct page (N × `LoadNextPage`)
4. Resume row selection

This is needed for **per-row detail extraction** (each row's General/Address/etc. tabs). The extension does not yet do per-row extraction (see § 15 roadmap); when ported, this recovery must come with it.

### 5.5 Known limitations

| Limitation | Workaround |
|---|---|
| ~160 KB response cap | Request specific tabs, not the whole form |
| 25-row default grid page | Always paginate; never trust first page |
| Limited fuzzy search in `form_find_menu_item` | Validate menu items with Playwright path validator |
| Segmented controls (financial dimensions) blank | Switch to Playwright for those forms |
| `form_find_controls` returns 3 duplicate copies per control; the 1st often empty | Don't rely on it for per-row values; use OData if needed |
| Token expiry on long batches | Re-auth via `az` between batches; keep batches ≤ 10 forms |

---

## 6. Playwright (Chrome CDP) Backend — Deep Dive

### 6.1 Connection

```python
playwright.chromium.connect_over_cdp("http://localhost:9222")
```

Prereq: `D365 FO Config Compare: Start Chrome with Remote Debugging` (or the legacy `start_chrome_cdp.bat`) — closes any running Chrome, relaunches with `--remote-debugging-port=9222 --user-data-dir=<profile>`. User signs in once; Playwright attaches.

### 6.2 Navigation

```
<env>/?cmp=<le>&mi=<menuItem>
```

`?cmp=` sets legal entity; `mi=` opens the form. Wait for the form chrome to render before scraping.

### 6.3 What gets scraped

| Element | Selector strategy |
|---|---|
| Tabs | `[role="tab"]` + the F&O FastTab CSS class set |
| Fields | `_JS_TABS` evaluator pulls `<input>` / combobox / checkbox label + value from each tab's panel |
| Grids | `_scroll_collect_grids` virtual-scrolls each container, captures cells as the DOM virtualises |
| Action buttons | `_ACTION_BUTTONS` allowlist — only safe activators |

### 6.4 Safety guards (mandatory)

- Hard allowlist of clickable selectors — tab expanders, scroll triggers only.
- All other interaction is read-style: focus, hover, scroll.
- The Python prompt + selectors **never** match Save / New / Delete / Edit / Post.
- Run only against UAT or read-replica environments; never production without dry-run.

### 6.5 When Playwright wins

- Forms with empty MCP tabs.
- Financial-dimension segmented controls.
- Visual verification ("show me what was scraped").
- Tenants where MCP is disabled.

### 6.6 Where Playwright loses

- Latency on slow networks (DOM round-trips).
- Headless impossible (needs a real Chrome window).
- Risk of accidental clicks if selectors broaden — keep the allowlist tight.

---

## 7. Form-Path Validator

Users provide UI paths in human-readable form:

```
Accounts payable > Setup > Vendor groups
```

The validator discovers the `mi=` menu item by driving Chrome via CDP:

```
1. Navigate to D365 home (clean slate)
2. Dismiss any open notification flyout
       (otherwise [role="listitem"] grabs toast items)
3. Click the magnifier (Search) button
       fallback chain: aria-label="Search for a page" → "Search"
       → click the SVG icon coords if is_visible() lies
4. Type the leaf segment of the path
5. Wait for results, filter by breadcrumb matching parent segments
6. Click the match; detect URL change (mi= captured)
7. Persist {path, menuItem, validated: true} to settings.json
```

### Validator gotchas (real bugs hit during this project)

| Bug | Fix |
|---|---|
| Notification flyout polluted result list | Dismiss `[role="listitem"]` with class containing "notification" before search |
| Magnifier `aria-label` differs by tenant locale | Fallback chain (`Search for a page` → `Search` → SVG selector) |
| `is_visible()` returns false on icon-only button | Click by coords if all fallbacks fail |
| `mi=` was captured as `DefaultDashboard` | Detect SPA URL change before reading; retry if `mi` didn't update |
| `?mi=NavigationSearchPage` deep-link doesn't exist in some tenants | Removed deep-link; always use UI search |
| UnicodeEncodeError on `✓` (cp1252 console) | All output ASCII-only (`OK`, `FAIL`) |
| envs were 4-tuple but unpacked as 3-tuple | `for l, _, _, _ in envs` |

---

## 8. Output Format

### Per-environment workbook

```
<form>_<le>_<env>_<timestamp>.xlsx
  Sheet: Form Info     caption, company, form name, timestamp
  Sheet: Top Fields    name | label | value | type | mandatory | visible
  Sheet: <TabName>     fields scoped to that tab (one sheet per tab)
  Sheet: <GridName>    paginated grid rows (one sheet per grid)
```

### Diff workbook (when 2+ environments)

```
<form>_<le>_DIFF_<timestamp>.xlsx
  Green     match
  Pink      values differ
  Yellow    present in one env, missing in another
```

Auto-filter is enabled. Frozen header row. Separator rows show LE + row counts: `MY30 - UAT: 47 rows | Config: 0 rows`.

### Why this layout

Functional consultants scan top-down. Field labels + values per tab = matches the F&O UI. Internal names go in a hidden column. The DIFF workbook is the **only** thing they need to action against — it's pre-filtered to mismatches.

---

## 9. Complete Gotcha Catalogue (20+ entries)

This is the single source of truth — supersedes [WORKAROUNDS_AND_LESSONS.md](archive/docs/WORKAROUNDS_AND_LESSONS.md).

### A. MCP parameter names (silent failures)

| Tool | Wrong | Correct |
|---|---|---|
| `form_find_controls` | `searchTerm` | `controlSearchTerm` |
| `form_select_grid_row` | `rowNumber: 1` (int) | `rowNumber: "1"` (string) |
| `form_open_or_close_tab` | no `tabAction` | `tabAction: "Open"` |
| `data_find_entities` | `entityName` | `odataPath` |
| `data_get_entity_metadata` | `entityName` | `entitySetName` |
| `data_find_entity_type` | `searchTerm` | `tableSearchFilter` |

### B. Grid pagination

- Max 25 rows per page.
- After `LoadNextPage`, grid data stays nested in `Tab → TabPage → Children → Grid` — search recursively, not at `Form > Grid > Name`.
- `HasNextPage` is a **string** `"True"/"False"`, not boolean.
- Even with `HasNextPage="False"` AND exactly 25 rows, try one more page — D365 lies occasionally.

### C. Tabs

- `form_open_or_close_tab` without `tabAction="Open"` silently returns ~9-12 fields instead of ~56-62 — biggest single source of "MCP missed fields."
- After opening a tab, **re-parse the FormState** — the open action returns fresh nested content.
- Some tabs contain sub-grids only present after open (e.g. "Transportation management" tab).

### D. `form_find_controls` returns triplicates

- Three copies of each control returned.
- 1st copy updates per row but often shows EMPTY for non-first rows.
- 2nd / 3rd copies retain STALE values from initial form load.
- **Don't use it for per-row FastTab values.** Use OData enrichment instead.

### E. Per-row state corruption

- After ~20-25 consecutive `form_select_grid_row` calls, MCP returns `Expecting value: line 1 column 1`.
- Recovery: close → reopen → paginate back N pages → resume.

### F. OAuth between legal entities

- Switching MY30 → MY60 → SG60 may trigger AAD timeouts.
- Retry: 3 attempts, 10s/20s/30s backoff.

### G. OData fuzzy match danger

- `data_find_entity_type` fuzzy-matches — picks **wrong entity** silently.
- Always verify: table match (`TableName_Field` control prefix), field metadata, value cross-check on a known record.

### H. Windows console encoding

- cp1252 chokes on `─ ═ ✓ →`.
- Wrap stdout: `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')` or write ASCII-only.

### I. D365 control naming

- Pattern is `TableName_FieldName` (e.g. `CustTable_CHBConsignmentWarehouse`).
- Use it to derive the backing table — more reliable than parsing form name.

### J. Multi-LE batching

- Sessions get flaky on long runs. Batch ≤ 10 forms.
- Output rate-limited writes; close form between runs.

### K. Validator robustness (see § 7)

All seven validator bugs listed above are gotcha entries.

### L. The "tab grids never merged" bug (cost hours)

In the original recursive parser:

```python
# WRONG — only fields merged, grids lost
data["fields"].update(tab_container["fields"])

# CORRECT
data["fields"].update(tab_container["fields"])
data["grids"].update(tab_container["grids"])
```

Both `Tab` and `TabPage` branches must merge `grids` too. The extension's `_parse_form` carries this fix.

### M. Cached form-path validation false positives

If a path was previously validated to the wrong `mi=` (e.g. `DefaultDashboard` from the SPA-navigation bug), it stays cached as `validated: true`. Fix: re-run `Configure Form Paths` and clear bad entries — or `Clear Saved Settings`.

### N. User-level vs workspace-level `defaultBackend`

A user-level setting of `playwright` suppresses the per-extraction backend picker. Either remove it or set it to `ask`.

### O. `form_open_menu_item` companyId is case-sensitive sometimes

Use canonical company codes from the company picker (`my30`, `usmf`) exactly as F&O returns them.

---

## 10. Authentication

### Extension model (current)

```
az account get-access-token --resource https://<env>.dynamics.com
→ bearer token
→ Authorization: Bearer <token> on /mcp request
```

- No app registration.
- No client secret.
- No persisted token — `az` is asked fresh each run.
- User must have read access to the env (granted via the same AAD account they use in D365).

### Original model (archive — `config.json`)

Pre-extension, the project used **client-credentials flow**:

```json
{
  "tenant_id": "00000000-0000-0000-0000-000000000000",
  "client_id": "00000000-0000-0000-0000-000000000001",
  "client_secret": "REDACTED"
}
```

This required an app registration with D365 access, stored secrets on disk, and had to be rotated. The extension removes all of that.

---

## 11. Original Python Pipeline (Pre-Extension Era)

The full Apr 28 – May 11 pipeline lives in [`archive/python/`](archive/python/). Annotated catalogue: see [`archive/README.md`](archive/README.md).

Key files (in order of importance):

1. **`form_control_extractor.py` (53.8 KB)** — the engine. 7-step pipeline. Ported to `vscode-ext/python/extract.py` as `extract_mcp`.
2. **`compare_form_controls.py`** — UAT vs Config diff → Excel. Ported to `write_excel` + `write_diff`.
3. **`deep_extract_parameters.py` (28 KB)** — parameters-form specialist (no main grid). Logic merged into stage 3 (force-open tabs) of the extension.
4. **`batch_extract_asset_fa.py` (40 KB)** — drove the 30+ Asset & FA Excel outputs.
5. **`form_reader.py` (34 KB)** — low-level MCP wrappers. Replaced by inline calls in `extract.py`.

### What didn't make the extension

- ADO ingestion (`fetch_all_ado_items.py`, `categorize_items.py`) — project-specific; would belong in a separate `ADO CDD Importer` extension if generalised.
- OData enrichment (`debug_custtable*`, `test_odata_*`) — needed only for CustTable's CBA/CHB FastTab fields. Out of scope for v1.
- 16 debug/test scripts — investigation tools, not user-facing.

---

## 12. VS Code Extension Architecture

### Files (in `vscode-ext/`)

| Path | Purpose |
|---|---|
| `package.json` | Extension manifest; commands, settings, walkthrough, icon |
| `src/extension.ts` | Activate, register all commands |
| `src/extractor.ts` | Multi-select form picker, env picker, backend picker, spawns Python |
| `src/onboarding.ts` | Python deps install, `az login` probe, env registration, `/mcp` probe |
| `src/validator.ts` | Orchestrates form-path validation |
| `src/welcome.ts` | Walkthrough wiring |
| `python/extract.py` | The actual extractor — both backends + validator + Excel writer |
| `walkthrough/01–05.md` | Step-by-step onboarding |
| `media/icon-beta.png` | Extension icon (gradient + diff cards + BETA ribbon) |
| `README.md` | User-facing quick start |
| `CHANGELOG.md` | Version log |
| `SKILL.md` | Concise domain knowledge |
| `MASTER_GUIDE.md` | This file |

### Commands (in execution order)

1. `D365 FO Config Compare: Getting Started` — opens walkthrough
2. `D365 FO Config Compare: Create / Open Workspace Folder` — scaffolds `.vscode/settings.json`
3. `D365 FO Config Compare: Run Onboarding` — deps + `az login` + envs + `/mcp` probe
4. `D365 FO Config Compare: Start Chrome with Remote Debugging` — `:9222` Chrome
5. `D365 FO Config Compare: Configure Form Paths` — paste UI paths, auto-discover `mi=`
6. `D365 FO Config Compare: Extract Form` — multi-select forms + envs + backend
7. `D365 FO Config Compare: Clear Saved Settings` — reset (rarely needed)

### Settings reference

| Setting | Type | Purpose |
|---|---|---|
| `d365FormExtractor.pythonPath` | string | Python executable (default `python`) |
| `d365FormExtractor.chromePath` | string | `chrome.exe` path (auto-detect) |
| `d365FormExtractor.environments` | array | `[{label, baseUrl, tenantId?, mcpUrl?, mcpAvailable?}]` |
| `d365FormExtractor.formPaths` | array | `[{path, menuItem, validated, error?}]` |
| `d365FormExtractor.outputFolder` | string | Where Excels are written |
| `d365FormExtractor.defaultBackend` | enum | `mcp` / `playwright` / `ask` |

---

## 13. Governance & Usage Model

### AI for read; humans for write

This entire toolchain is **read-only**. The extension never writes to D365. All proposed fixes flow through controlled channels:

- Functional consultant reviews the DIFF Excel.
- Decides which value is correct (Golden? UAT? a new third value?).
- Files an ADO ticket against the correct environment.
- A change is applied via the team's deployment pipeline, manual config, or DMF data upload — never by the extension.

### Why this boundary matters

- D365 production data has audit/compliance implications.
- Bulk automated changes risk cascading failures.
- The diff is the *evidence*; the human is the *decision*.

### What requires explicit user confirmation (in the extension)

- Closing Chrome (when launching CDP mode) — user clicks through.
- Overwriting an Excel that already exists — `outputFolder` timestamping prevents collisions.
- Clearing saved settings — separate command, not bundled.

---

## 14. Operational Runbook

### First-time setup

1. Install the VSIX or from Marketplace (when published).
2. `D365 FO Config Compare: Getting Started`.
3. Run walkthrough steps 1-5 in order.

### Adding a new environment

1. Open `.vscode/settings.json`.
2. Append to `d365FormExtractor.environments`:
   ```json
   { "label": "New Env", "baseUrl": "https://new-env.dynamics.com" }
   ```
3. Run `D365 FO Config Compare: Run Onboarding` again — it probes the new env's `/mcp`.

### Adding a new form

1. Open D365 in CDP Chrome, navigate to the form once.
2. `D365 FO Config Compare: Configure Form Paths` → paste the UI nav path.
3. Validator discovers `mi=`, persists it. Done.

### Running a comparison

1. `D365 FO Config Compare: Extract Form`.
2. Select legal entity.
3. Multi-select forms (Space to toggle, Enter to confirm).
4. Multi-select envs (≥ 2 to get a DIFF file).
5. Pick backend (MCP recommended, Playwright for complex forms).
6. Watch the Output panel.

### Troubleshooting decision tree

```
"Script hangs"
  → set PYTHONUNBUFFERED=1, retry

"Expecting value: line 1 column 1" (MCP)
  → Server state corrupt — auto-recovers; if repeats, env under load

"401 Unauthorized"
  → az login; check tenant; check D365 access

"Form shows 0 rows"
  → Normal for Config envs. Field-level diff still meaningful.

"Wrong values in Excel"
  → Check backend used. For CustTable FastTab fields → OData (not yet in ext)

"Menu item not found"
  → Check formPaths cache; try mi_type=Action; verify in D365 manually

"Validator never finds path"
  → Notification flyout open? Reload home. Verify breadcrumb wording.

"Tab shows 0 fields"
  → Pre-0.7.0 extension. Update to 0.8.x.

"Sweep didn't add any fields"
  → That's fine — every field was already found via tab walk.
```

---

## 15. Roadmap

### Shipped (current — 0.8.2 Beta)
- Validator hardened (notification, magnifier, breadcrumb, MI-change)
- Multi-select form extraction
- Deep MCP extraction (5-stage)
- Pagination, FastTab expansion, hidden-control sweep
- Multi-env DIFF workbook
- Walkthrough + icon + BETA marking
- Public GitHub repo + SKILL.md + this guide

### Next (high value)
1. **Port `_extract_row_details`** — per-row tab walk via `form_select_grid_row`. Includes the 25-row close/reopen recovery. Adds ~1-3 min/form but captures sub-grid detail fields (the CustTable CBA/CHB case).
2. **OData enrichment fallback** — for forms where MCP returns stale row values.
3. **Smoke test command** — pings each env's `/mcp`, opens a known-good form (`LedgerParameters`), prints per-tab field counts. Catches tenant-level MCP issues immediately.

### Later (nice to have)
4. **Marketplace publish** — `vsce publish` with PAT. Replaces VSIX file-passing.
5. **HTML diff report** alongside Excel — for sharing in PRs/Teams.
6. **Narrower sweep** — domain-aware term sets per form type (cuts MCP latency).
7. **ADO importer command** — generalised `fetch_all_ado_items.py` as a command. Auto-populate `formPaths` from a CDD query.
8. **Per-extraction telemetry** — record extraction time, sweep adds, paginations to a local SQLite. Helps tune the sweep term set.

### Won't do
- Write operations against D365. Read-only is the governance boundary.
- Headless Playwright. Defeats the auth model.
- Bundling Chrome. Use the user's installed Chrome.

---

## 16. File Reference (Old + New)

### Active (the extension)
| File | Purpose |
|---|---|
| `vscode-ext/package.json` | Manifest |
| `vscode-ext/python/extract.py` | Both backends + validator + Excel |
| `vscode-ext/src/*.ts` | Host code |
| `vscode-ext/walkthrough/0[1-5]-*.md` | Onboarding |
| `vscode-ext/README.md` | Quick start |
| `vscode-ext/SKILL.md` | Domain knowledge (concise) |
| `vscode-ext/MASTER_GUIDE.md` | This file |
| `vscode-ext/CHANGELOG.md` | Version log |

### Archive (historical context)
| Folder | Content |
|---|---|
| [`archive/python/`](archive/python/) | 57 Python scripts + `start_chrome_cdp.bat` |
| [`archive/docs/`](archive/docs/) | TEAM_GUIDE (md/docx/pdf), WORKAROUNDS_AND_LESSONS, MCP vs Playwright comparison docx |
| [`archive/ado-data/`](archive/ado-data/) | 5 JSON files — raw ADO export + categorisation |
| [`archive/agent-config/`](archive/agent-config/) | Original Claude/Copilot agent definition |
| [`archive/local-replica/`](archive/local-replica/) | Self-contained Apr 29 snapshot |

### External references (not in repo)
- `C:\D365DataValidator\config.json` — original env config (client-credentials). Obsolete; extension uses `az`.
- `C:\D365DataValidator\d365_mcp_client.py` — original MCP client. Replaced by inline calls.
- `C:\Bugs daily report\` — sibling project (ADO 6 AM SGT email automation). Separate concern.

---

## End

This document is meant to be **the** reference. If you find yourself asking "where do I read about X?", the answer should be findable from the TOC above. If it isn't, that's a bug — open an issue.

**Repo:** https://github.com/prashantdce21MSFT/d365-fo-config-compare  
**Version of this document:** 1.0 — 2026-05-13  
**Project span:** 2026-04-28 → 2026-05-13 (16 days, ~70 scripts, 30+ Excel outputs, 1 published extension)
