# Archive — D365 FO Config Drift Analysis Project

This folder is a **frozen historical record** of the project that led to the `D365 FO Config Compare` VS Code extension. Nothing here is needed to run the extension. It is preserved so the team can:

- understand *how* and *why* the extension was built,
- read the original gotchas/lessons in their original form,
- recover any Python prototype if needed,
- and trace the lineage from "606 CDDs in ADO" → "an installable VS Code extension."

> **The extension** lives in the parent folder (`../`). Use it for actual work.
> **This archive** is for context.

---

## Folder Map

| Folder | What's inside | Snapshot date |
|---|---|---|
| [`python/`](python) | All 57 Python scripts + `start_chrome_cdp.bat` — the original Acme REGION analysis pipeline | Apr 28 – May 11 2026 |
| [`docs/`](docs) | Team guide, lessons learned, MCP-vs-Playwright comparison (.md / .docx / .pdf) | Apr 29 – May 11 2026 |
| [`ado-data/`](ado-data) | Raw ADO export (606 CDD items, categorisation, checkpoint) | Apr 28 2026 |
| [`agent-config/`](agent-config) | Original `.github/agents/d365-drift-analysis.agent.md` (Claude/Copilot agent definition) | May 1 2026 |
| [`local-replica/`](local-replica) | Self-contained "go-anywhere" snapshot from Apr 29 (core + shared_lib + cache) | Apr 29 2026 |

---

## What to read first

| If you want to… | Read |
|---|---|
| Understand the business problem & full E2E flow | [`docs/TEAM_GUIDE_D365_Config_Drift_Analysis.md`](docs/TEAM_GUIDE_D365_Config_Drift_Analysis.md) |
| See every MCP/Playwright gotcha we hit | [`docs/WORKAROUNDS_AND_LESSONS.md`](docs/WORKAROUNDS_AND_LESSONS.md) |
| Read the formal MCP-vs-Playwright comparison | [`docs/D365_ERP_MCP_vs_Playwright_Comparison.docx`](docs/D365_ERP_MCP_vs_Playwright_Comparison.docx) |
| Read what the deep extractor does (the code that became the extension) | [`python/form_control_extractor.py`](python/form_control_extractor.py) |

---

## Python script catalogue

Grouped by purpose. All in [`python/`](python).

### ADO ingestion (one-time, Apr 28)
| File | Purpose |
|---|---|
| `fetch_all_ado_items.py` | Pull 606 CDD work items from ADO query `7a10e4a5-...` |
| `categorize_items.py` | Score-match each CDD title → D365 menu item |
| `generate_cdd_excel.py` | Parse `[MY]/[SG]` tags + nav paths → `CDD_Form_Paths_*.xlsx` |
| `download_screenshots.py` | Download attachments from each work item |
| `mysg_reporter.py` | Aggregate ADO report builder |

### Form discovery / mapping (Apr 28 – May 1)
| File | Purpose |
|---|---|
| `find_correct_mi.py` | URL-scrape `mi=` from D365 navigation |
| `verify_mi_urls.py` | Smoke-test each menu item URL |
| `build_form_mapping.py` | Map UI nav-paths to menu items (53 KB — the big one) |
| `build_tables_mapping.py` | Map forms to backing tables |
| `merge_cdd_tables.py` | Merge CDD list with form-table map |
| `read_form_all_companies.py` | Same form across MY30 / MY60 / SG60 |
| `read_invoice_charge_tracking.py` | One-off form read |

### Extraction engines (Apr 28 – Apr 30) — **the lineage of the extension**
| File | Purpose |
|---|---|
| **`form_control_extractor.py`** | **53.8 KB. The original deep extractor.** Ported into `vscode-ext/python/extract.py` as `extract_mcp`. 5-stage pipeline (parse → expand FastTabs → force-open tabs → paginate grids → ~146-term sweep) lives here in its original form. |
| `form_reader.py` | Lower-level MCP wrappers (`open_form`, `close_form`, `mcp_call`, `resolve_menu_item_from_path`) |
| `deep_extract_parameters.py` | Parameters-form specialist (no main grid, all tabs) |
| `run.py` | One-off runner |
| `run_batch_compare.py` | Multi-form sequential driver |

### Batch runners (Apr 28 – May 4)
| File | Purpose |
|---|---|
| `batch_all_form_compare.py`, `batch_compare_v2.py` | Multi-form comparison, v1 & v2 |
| `batch_extract_config.py` | Mass-extract from Env4 Config env |
| `batch_extract_asset_fa.py` | 40 KB — drives the 30+ Asset/FA Excel outputs |
| `batch_find_tables.py` | Discover tables per form |
| `batch_form_compare.py` | Final shape used May 11 |
| `batch_odata_compare.py` | OData-only fallback comparator |

### Comparison + reporting (Apr 28 – May 4)
| File | Purpose |
|---|---|
| `compare_form_controls.py` | Per-form UAT vs Config → Excel |
| `compare_61057.py`, `compare_order_purposes.py`, `compare_single_form.py` | Targeted comparisons |
| `build_comparison.py`, `build_final_report.py` | Final-report builders |
| `merge_results_to_excel.py` | Stitch multi-LE JSON → workbook |
| `get_row_counts.py` | Grid-row counts per env/LE |
| `check_custtable_excel.py` | Smoke-check CustTable extraction |

### Debug / one-off tests (Apr 29 – Apr 30)
| File | Purpose |
|---|---|
| `debug_find_controls*.py` (3) | Investigated the 3-duplicate-copy bug |
| `debug_custtable*.py` (3) | CustTable per-row stale-value investigation |
| `debug_interest*.py` (3) | Interest-form value debugging |
| `debug_returnaction*.py` (2) | ReturnAction state bug |
| `debug_ra_cols.py` | Column dump helper |
| `test_consignment_values.py`, `test_custtable_fasttab.py`, `test_dedup.py`, `test_fasttab_click.py`, `test_fc_per_row.py`, `test_find_tables_3forms.py`, `test_odata_custtable*.py` (2), `test_odata_cba_fields.py`, `test_one.py` | Misc smoke tests |
| `sample_36293.py` | Reproducer for one ADO item |

### Utilities
| File | Purpose |
|---|---|
| `md_to_docx.py` | Pandoc-free MD → DOCX converter |
| `generate_cdd_excel.py` | CDD → Excel (also in ADO section) |
| `start_chrome_cdp.bat` | Launches Chrome with `--remote-debugging-port=9222` — replaced by the extension's `Start Chrome with Remote Debugging` command |

---

## ADO data

| File | Content |
|---|---|
| `ado-data/ado_query_raw.json` | 840 KB — raw response from the ADO query API |
| `ado-data/ado_items.json` | 130 KB — flattened work-item list |
| `ado-data/categorization_v2.json` | 244 KB — categorisation results |
| `ado-data/cat_checkpoint.json` | 298 KB — categoriser progress checkpoint |
| `ado-data/categorization_result.json` | 24 KB — final categorisation pass |

These are project-specific to Acme MY/SG. They're here for completeness; you would not commit equivalents from a fresh customer engagement to a public repo.

---

## Why nothing was deleted

The user explicitly asked that this archive be additive, never destructive. The original files still live at `C:\D365 Configuration Drift Analysis\` on the author's workstation. This archive in the repo is a **copy** to make the history portable and reviewable on GitHub.

---

## Mapping: archive → extension

If you wonder where a piece of the old pipeline ended up in the productised extension:

| Old (archive) | New (extension) |
|---|---|
| `python/form_control_extractor.py` → `_extract_for_le`, `_parse_form_obj`, `_paginate_grid`, `FIND_CONTROLS_TERMS`, sweep | `vscode-ext/python/extract.py` → `extract_mcp`, `_parse_form`, `_paginate_grid`, `_find_controls_sweep` |
| `python/start_chrome_cdp.bat` | Command: `D365 FO Config Compare: Start Chrome with Remote Debugging` |
| `python/find_correct_mi.py`, `verify_mi_urls.py` | Command: `D365 FO Config Compare: Configure Form Paths` + `validate_form_paths` in `extract.py` |
| `python/compare_form_controls.py` + `build_comparison.py` | `write_excel` + `write_diff` in `vscode-ext/python/extract.py` |
| `python/batch_*` runners | Multi-select in `vscode-ext/src/extractor.ts` |
| `C:\D365DataValidator\config.json` (client_secret-based) | `d365FormExtractor.environments` setting + Azure CLI bearer token (no secret stored) |
| `docs/WORKAROUNDS_AND_LESSONS.md` (15 gotchas) | Distilled into `vscode-ext/SKILL.md` § 8 (key takeaways) |
| `docs/TEAM_GUIDE_*.md` (run-by-hand workflow) | `vscode-ext/README.md` + `vscode-ext/walkthrough/01–05.md` |
| `agent-config/d365-drift-analysis.agent.md` | (No equivalent yet — see roadmap) |

---

**Snapshot:** 2026-05-13. Anything dated later belongs in the extension, not here.
