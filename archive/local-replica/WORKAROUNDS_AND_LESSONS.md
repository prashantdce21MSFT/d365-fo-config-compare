# FormControlExtractor - Workarounds & Lessons Learned

## MCP Form Tool Workarounds

### 1. `form_find_controls` parameter name
- **Wrong**: `searchTerm`
- **Correct**: `controlSearchTerm`
- The MCP tool uses `controlSearchTerm` as the parameter name for searching controls.

### 2. `form_find_controls` response format
- Returns a JSON **list** of `{Name, Properties}` objects, NOT a dict with a `Controls` key.
- Empty response = empty string (not valid JSON) — always handle gracefully with try/except.

### 3. `form_select_grid_row` requires string rowNumber
- `rowNumber` must be passed as a **string**, not an integer: `{"rowNumber": "0"}` not `{"rowNumber": 0}`.

### 4. Grid pagination (25-row pages)
- D365 MCP returns max 25 rows per grid page.
- Use `form_click_control` with `actionId: "LoadNextPage"` to get the next page.
- Check `Pagination.HasNextPage` (string `"True"/"False"`) to know when to stop.
- **Critical**: After `LoadNextPage`, the grid data in the response is still nested inside the Tab/TabPage hierarchy — you must search recursively, not just at `Form > Grid > gridName`.

### 5. Grids nested inside Tabs
- Grids can be at `Form > Tab > TabName > Children > TabPage > Children > Grid > GridName`.
- The `_parse_form_obj` method must recurse into Tab and TabPage children.
- **Bug that cost hours**: After recursing into tab children, `tab_container["grids"]` was populated but never merged into `data["grids"]`. Only `data["fields"]` was merged. Fix: add `data["grids"].update(tab_container["grids"])` in both Tab and TabPage processing blocks.
- Same applies to `_paginate_grid` — the LoadNextPage response keeps grids nested inside Tabs. Use `_find_grid_in_form()` to recursively locate the grid by name.

### 6. Tab sub-grids (Transportation management, DimensionIntegration)
- Some tabs contain sub-grids that only appear after opening the tab.
- Use `form_open_or_close_tab` first, then re-read FormState or run `form_find_controls`.
- Parenthetical hints in tab labels help with matching (e.g., "General (transportation)").

---

## Windows / Python Workarounds

### 7. Unicode encoding errors (cp1252)
- Windows console uses cp1252 encoding by default.
- Box-drawing characters (`─`, `═`) and em-dashes (`—`) cause `UnicodeEncodeError`.
- **Fix**: Replace all box-drawing chars with ASCII equivalents (`-`, `=`) in print statements.

### 8. OAuth connection timeouts between Legal Entities
- When creating a new MCP client for the 2nd/3rd legal entity, `login.microsoftonline.com` can timeout.
- **Fix**: Add retry logic (3 attempts, 10s/20s backoff) around `client.connect()` in `_extract_for_le`.

---

## Excel Output Lessons

### 9. Grid data presentation
- **Wrong approach**: Breaking grid cells into individual rows (`GridName[0].Formula`, `GridName[0].Name`) — creates thousands of unreadable rows.
- **Correct approach**: Present grid data as a proper D365-style table with actual column headers (Formula, Name, Site, Item group, etc.) — one row per record, just like the Standard view in F&O.
- Grid data and field-level controls need **separate sheets** because they're fundamentally different data types.

### 10. Comparison layout for grids
- Group by Legal Entity with separator rows showing row counts: "MY30 - UAT: 1551 rows | Config: 0 rows"
- UAT rows first, then Config rows for each LE.
- Add `Env` and `Legal Entity` columns for filtering.
- Use auto-filter so users can filter by environment or LE.

### 11. Fields sheet must include row details with Grid Row identifier
- **Wrong approach**: Skipping `].detail.` entries from Fields sheet — hides per-row field values entirely, so mismatches in row-specific fields (e.g., Site, Warehouse type) are invisible.
- **Wrong approach**: Showing row detail fields without identifying which grid row they belong to — "Site = 61" is meaningless without knowing it's for warehouse "61 - Duty Free".
- **Correct approach**: Include row detail fields in Fields sheet with a **Grid Row** column showing a readable identifier (first 1-2 column values from the grid row, e.g., "61 - Duty Free").
- Row details come from `form_select_grid_row` — selecting each grid row reveals detail fields (General tab, sub-tabs, etc.) that aren't in the grid columns.
- The `merge_le_controls` function must attach `grid_row_id` and `grid_name` to each row detail entry.

---

## Architecture Lessons

### 12. Separation of extraction and comparison
- `FormControlExtractor` class handles extraction only — give it MI name, LEs, environment.
- `compare_form_controls.py` handles comparison — runs extractor on two environments and diffs.
- This separation lets you extract once and compare multiple times, or compare different environment pairs.

### 13. Data reality across environments
- UAT environments have real transactional/master data (formulas, BOMs, etc.).
- Config/Golden environments often have zero records for data-heavy forms.
- Not all LEs have data — MY60 and SG60 may have no formula records while MY30 has 1551.
- Comparison is only meaningful where data exists.

### 14. Control name patterns
- D365 control names follow `TableName_FieldName` pattern (e.g., `BOMTable_Name`, `CustInterest_InterestCode`).
- Use this pattern to extract table names from controls — more reliable than guessing from form name.
- Grid controls have a separate naming pattern — the grid name itself (e.g., `TableGrpGrid`) contains the rows.

### 15. form_find_controls sweep strategy
- Use broad search terms (a-z, common prefixes) to discover controls not visible in initial FormState.
- Deduplicate against already-captured fields and grids.
- Skip `SystemDefined*` controls for cleaner output.

---

## Key Files

| File | Purpose |
|------|---------|
| `form_control_extractor.py` | Reusable class — opens form, parses FormState, paginates grids, extracts all controls |
| `compare_form_controls.py` | Comparison script — runs extractor on ENV1 vs ENV4, generates D365-style Excel |
| `C:\D365DataValidator\d365_mcp_client.py` | MCP client — `D365McpClient`, `call_tool()`, `connect()` |
| `C:\D365DataValidator\config.json` | Environment configs (ENV1=Env1 UAT, ENV4=Env4 Config) |
| `output\batch_extract_mi_cache.json` | Navigation path to menu item name mapping |

## Environment Reference

| Key | Name | URL |
|-----|------|-----|
| ENV1 | Env1 UAT | example-uat.sandbox.operations.dynamics.com |
| ENV2 | Env2 Golden | example-goldconfig.sandbox.operations.eu.dynamics.com |
| ENV3 | Env3 UAT | example-uat-eu.sandbox.operations.eu.dynamics.com |
| ENV4 | Env4 Config | example-config.sandbox.operations.dynamics.com |
