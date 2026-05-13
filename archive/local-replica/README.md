# D365 Form Control Extractor & Comparator - Local Replica

Self-contained copy of all files needed for the FormControlExtractor tool.
This is a snapshot so future conversations have everything in one place.

## Folder Structure

```
_local_replica/
  WORKAROUNDS_AND_LESSONS.md   -- All workarounds, bugs, and lessons learned
  README.md                     -- This file
  core/
    form_control_extractor.py   -- Main class: FormControlExtractor
    compare_form_controls.py    -- UAT vs Config comparison + Excel output
    form_reader.py              -- Lower-level MCP form helpers (open, close, mcp_call, etc.)
  shared_lib/
    d365_mcp_client.py          -- D365McpClient (MCP connection, call_tool, connect)
    config.json                 -- Environment configs (ENV1-ENV4, credentials, URLs)
  cache/
    batch_extract_mi_cache.json -- Navigation path -> menu item name mapping
```

## How It Works

### 1. FormControlExtractor (core/form_control_extractor.py)

```python
from form_control_extractor import FormControlExtractor

ext = FormControlExtractor(env_key="ENV1")  # ENV1=Env1 UAT, ENV4=Env4 Config
result = ext.extract("PmfFormulaTable", ["MY30", "MY60", "SG60"], "Display")
```

**What it does per legal entity:**
1. Connects to MCP server via D365McpClient
2. Opens form via `form_open_menu_item`
3. Parses FormState recursively (grids, fields, tabs, buttons)
4. Paginates grids (25 rows/page) via `form_click_control` + `LoadNextPage`
5. Selects each grid row to get detail fields
6. Opens tabs and re-reads controls
7. Runs `form_find_controls` sweep (a-z + common terms)
8. Closes form

**Returns:** Nested dict with grids (rows, columns, pagination), fields, tabs, buttons, summary per LE.

### 2. Comparison (core/compare_form_controls.py)

```python
python compare_form_controls.py
```

- Runs FormControlExtractor on ENV1 (UAT) and ENV4 (Config)
- Builds comparison across all LEs
- Generates Excel with:
  - **Summary** sheet: form info, per-LE row counts
  - **Grid sheets**: D365-style tables (one row per record, actual column headers)
  - **Fields** sheet: non-grid control comparison with Match/MISMATCH

### 3. Dependencies

- `form_reader.py` — `mcp_call()`, `open_form()`, `close_form()`, `resolve_menu_item_from_path()`
- `d365_mcp_client.py` — `D365McpClient` class (OAuth, MCP JSON-RPC)
- `config.json` — environment URLs and credentials
- `openpyxl` — Excel generation (pip install openpyxl)

## Key Workarounds

See `WORKAROUNDS_AND_LESSONS.md` for the full list. The critical ones:

1. **Grid pagination recursive search** — After LoadNextPage, grid data stays nested inside Tabs. Use `_find_grid_in_form()` to search recursively.
2. **Tab grids not merging** — Must add `data["grids"].update(tab_container["grids"])` after recursing into Tab/TabPage children.
3. **`controlSearchTerm`** not `searchTerm` for `form_find_controls`.
4. **OAuth retry** — 3 attempts with 10s/20s backoff for connection timeouts between LEs.
5. **cp1252 encoding** — No Unicode box-drawing chars in print statements on Windows.

## Environment Reference

| Key  | Name          | URL                                                    |
|------|---------------|--------------------------------------------------------|
| ENV1 | Env1 UAT        | example-uat.sandbox.operations.dynamics.com        |
| ENV2 | Env2 Golden | example-goldconfig.sandbox.operations.eu.dynamics.com    |
| ENV3 | Env3 UAT   | example-uat-eu.sandbox.operations.eu.dynamics.com           |
| ENV4 | Env4 Config   | example-config.sandbox.operations.dynamics.com     |

## Snapshot Date
2026-04-29
