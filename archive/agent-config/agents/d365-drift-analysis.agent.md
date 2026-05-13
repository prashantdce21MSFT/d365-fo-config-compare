---
name: D365 Drift Analysis Agent
description: "Use when: performing D365 form extraction, comparing configurations between UAT and Config environments, analyzing control values, handling form tabs and sub-grids, or troubleshooting extraction issues"
keywords: ["D365", "form extraction", "configuration drift", "UAT vs Config", "smmParameters", "tabs", "pagination"]
---

# D365 Configuration Drift Analysis Agent

**Project**: Acme MY/SG Configuration Drift (UAT `example-uat` vs Config `example-config`)  
**Team**: Acme Asia QA on SmartCore ASIA  
**Shared with**: Team collaboration on drift analysis across MY30, MY60, SG60 legal entities

## Project Memory & Context

This agent consolidates **all team knowledge** about this solution:

### User Profile
- **Prashant Verma**, Acme Asia QA on SmartCore ASIA
- Use `Acme.asia` account (NOT personal Azure tenant)
- [Full profile](../../../Users/prverma/.claude/projects/c--Acme-SQL-Analaysis-for-DM/memory/user_profile.md)

### Project Reference
- **MY/SG Config Drift**: UAT vs Env2 Golden across 3 LEs (MY30, MY60, SG60)
- **ADO Query**: `7a10e4a5-...` (606 items to track)
- **Project Dir**: `C:\D365 Configuration Drift Analysis\`
- **Shared Lib**: `C:\D365DataValidator\` (FormControlExtractor, helpers)
- [Full reference](../../../Users/prverma/.claude/projects/c--Acme-SQL-Analaysis-for-DM/memory/project_mysg_drift.md)

### Extraction Patterns & Gotchas

#### MCP Form Tool Essentials
- **tabAction required**: `form_open_or_close_tab(tabName, tabAction="Open")` — silently fails without it
- **Encoding issue**: cp1252 codec can't encode Unicode arrows (→); fix with `io.TextIOWrapper` UTF-8 wrapper
- **Parameter substitution**: use `controlSearchTerm` (not `searchTerm`), `rowNumber` as string
- **Grid nesting**: Grids appear nested in Tabs after LoadNextPage; use recursive `find_grid_in_form()`
- **OAuth retry**: Between legal entities, may need reconnect with backoff
- [Detailed gotchas](../../../Users/prverma/.claude/projects/c--Acme-SQL-Analaysis-for-DM/memory/feedback_mcp_form_gotchas.md)

#### Form Extraction Pipeline
1. **Open form** → Parse FormState.Form recursively
2. **Grid pagination** → LoadNextPage until `HasNextPage=false` (25-row pages)
3. **Row details** → Open sub-tabs per row, extract tab-specific fields
4. **Tab walking** → Open each form-level tab, re-parse to discover hidden controls
5. **find_controls sweep** → 63-term broad search (after all tabs open)
6. **Sub-grid pagination** → Paginate sub-grids within tabs
7. **Merge & deduplicate** → Tab-sourced fields win over sweep results

**Key insight**: `form_open_or_close_tab` returns full FormState — re-parsing reveals controls hidden when tab was collapsed.

#### Parameters Forms (No Grids)
- All data lives in tab-level fields, not a main grid
- find_controls sweep alone is insufficient (e.g., smmParameters had only 72 controls)
- Deep extraction (walk all tabs + re-parse) yields 52% more coverage
- Use `deep_extract_parameters.py` pattern for parameters forms
- [Parameters form guidance](../../../Users/prverma/.claude/projects/c--Acme-SQL-Analaysis-for-DM/memory/feedback_parameters_forms.md)

#### Form Table Extraction
- Extract table names from **control name prefixes** (`TableName_FieldName`), not form name
- Use `form_find_controls` with broad searches, then parse Properties
- [Extraction approach](../../../Users/prverma/.claude/projects/c--Acme-SQL-Analaysis-for-DM/memory/feedback_form_table_extraction.md)

#### Excel Output Layout
- **Summary sheet** → Form metadata, env/LE info, status
- **Grid sheets** (one per grid) → D365-style tables with LE separator rows and counts
- **Fields sheet** → Non-grid controls from find_controls + form fields
- **Sub-grid sheets** → Separate tables for paginated rows within tabs
- Auto-filter, frozen panes, color-coded match status
- [Layout standard](../../../Users/prverma/.claude/projects/c--Acme-SQL-Analaysis-for-DM/memory/feedback_grid_excel_layout.md)

### Form Tools & Pagination
- **25-row page limit** → All grid types (forms, sub-grids, sub-tabs)
- **LoadNextPage action** → Required for each subsequent page
- **Tab sub-grids** → May exist in Transportation mgmt, DimensionIntegration, others
- **Parenthetical hints** → Use hints like `(Transportation mgmt)` to match nested tabs
- [Pagination patterns](../../../Users/prverma/.claude/projects/c--Acme-SQL-Analaysis-for-DM/memory/feedback_form_tools_pagination.md)

### Code Organization
- **FormControlExtractor** (shared library) → 7-step extraction pipeline
- **form_control_extractor.py** → Grid-focused extraction + comparison Excel
- **deep_extract_parameters.py** → Parameters form deep extraction (all tabs + sub-grids)
- **form_reader.py** → MCP helpers (open_form, close_form, mcp_call)
- **compare_form_controls.py** → Comparison + Excel generation
- **_local_replica/** → Self-contained snapshot with 15 workarounds
- [Architecture reference](../../../Users/prverma/.claude/projects/c--Acme-SQL-Analaysis-for-DM/memory/ref_local_replica.md)

### OData Enrichment (CustTable Only)
- Acme Asia FastTab fields (CBA/CHB prefixed) available via OData CustomersV3
- Use for enriching grid row details when form values are stale
- Always verify backing table, field names before trusting OData results
- [OData cautions](../../../Users/prverma/.claude/projects/c--Acme-SQL-Analaysis-for-DM/memory/feedback_odata_entity_matching.md)

### MCP Architecture
- **No Claude MCP add** for corporate servers (Dataverse, ADO, Power Platform)
- Build **standalone CLIs** that subprocess the MCP server instead
- Use `D365McpClient` (subprocesses MCP server internally)
- [Architecture decision](../../../Users/prverma/.claude/projects/c--Acme-SQL-Analaysis-for-DM/memory/feedback_no_claude_mcp.md)

## Current Workflows

### Deep Extraction (Parameters Forms)
**Script**: `deep_extract_parameters.py`  
**Use for**: Forms with many tabs but no main grid (e.g., smmParameters)

```python
# Opens each tab, re-parses FormState, paginates sub-grids
extract_env(mi_name="smmParameters", mi_type="Display", 
            legal_entities=["MY30", "MY60", "SG60"],
            env_key="ENV1", env_label="Env1 UAT")
```

**Output**: One sheet per tab, with fields grouped by tab + sub-grid rows

### Grid Extraction (Form-Based)
**Script**: `form_control_extractor.py` + `compare_form_controls.py`  
**Use for**: Forms with main grid(s) and row-detail tabs (e.g., InventLocations, CustTable)

**Run**:
```python
extractor = FormControlExtractor(env_key="ENV1")
result = extractor.extract(mi_name="...", legal_entities=["MY30", "MY60", "SG60"])
```

## Quick Checklists

### Before Running Extraction
- [ ] Confirm MI name (use MI cache or URL screenshot)
- [ ] Verify form type (Display, Action, Output)
- [ ] Check legal entities (MY30, MY60, SG60)
- [ ] Set env_key (ENV1 for UAT, ENV4 for Config)
- [ ] Fix encoding: wrap stdout with UTF-8 for Unicode support
- [ ] Ensure D365DataValidator lib is in sys.path

### Debugging Failed Extractions
- [ ] Check connection (3x retry with backoff)
- [ ] Verify tabAction="Open" on tab opens
- [ ] Inspect FormState response for tab-specific fields
- [ ] Paginate grids with LoadNextPage until HasNextPage=false
- [ ] Run find_controls sweep AFTER all tabs opened
- [ ] Check Excel for LE separator rows and field counts

### Verification Steps
1. Run single LE (MY30) on one env first (fast feedback)
2. Verify tab count matches what user sees in D365
3. Check each tab sheet has fields with values
4. Verify sub-grids have all rows (not truncated to 20)
5. Run full 3 LEs × 2 envs
6. Spot-check known values between UAT and Config

## Related Projects

- **Region3 Validator** (`C:\D365Region3Validator\`) — Template for config drift analysis
- **Daily ADO Bug Reports** (`C:\Bugs daily report\`, `C:\Bugs Daily report HUB\`) — Scheduled extraction tasks using similar patterns
- **DataverseCompare** (`C:\DataverseCompare\`) — Dataverse entity comparisons

---

**Last updated**: May 1, 2026  
**Agent owner**: Team (shared in `.github/agents/`)
