"""
Form Reader — reads data from D365 forms via MCP form tools.
Used for entities not exposed via OData publicly.
"""
import json
import time
import re


def mcp_call(client, tool_name, args):
    """Call an MCP tool and return parsed result."""
    result = client.call_tool(tool_name, args)
    if isinstance(result, str):
        try:
            return json.loads(result)
        except json.JSONDecodeError:
            return {"raw": result}
    return result


def find_menu_item(client, search_term, company_id="HR01"):
    """Search for a menu item by name. Returns list of matches."""
    result = mcp_call(client, "form_find_menu_item", {
        "companyId": company_id,
        "menuItemFilter": search_term,
    })
    menu = result.get("Menu", {})
    items = []
    for item_type in ("Display", "Output", "Action"):
        for item in menu.get("MenuItems", {}).get(item_type, []):
            items.append({
                "name": item.get("Name", ""),
                "text": item.get("Text", ""),
                "type": item_type,
            })
    return items


def open_form(client, menu_item_name, menu_item_type="Display", company_id="HR01"):
    """Open a D365 form and return parsed form state."""
    result = mcp_call(client, "form_open_menu_item", {
        "name": menu_item_name,
        "type": menu_item_type,
        "companyId": company_id,
    })
    if "raw" in result:
        raw = result.get("raw", "")
        if "exception" in raw.lower() or "error" in raw.lower():
            raise RuntimeError(f"Failed to open form: {raw}")
    return result


def close_form(client):
    """Close all open forms."""
    try:
        mcp_call(client, "form_close_form", {})
    except Exception:
        pass


def extract_form_data(client, form_result):
    """
    Extract all data from an opened form.

    Returns dict with:
      - grids: {grid_name: {columns: [...], rows: [{values}, ...]}}
      - fields: {field_name: {label, value, type}}  (from Input/Checkbox/Combobox)
      - tabs: {tab_name: {text, fields: {field_name: {label, value, type}}}}
      - records: list of dicts (combined grid + detail per row)
    """
    form_state = form_result.get("FormState", {})
    form = form_state.get("Form", {})

    data = {
        "form_name": form_state.get("Name", ""),
        "caption": form_state.get("Caption", ""),
        "company": form_state.get("Company", ""),
        "grids": {},
        "fields": {},
        "tabs": {},
        "records": [],
    }

    # Extract grids
    for grid_name, grid_info in form.get("Grid", {}).items():
        if not isinstance(grid_info, dict):
            continue
        columns = [c.get("Label", c.get("Name", "")) for c in grid_info.get("Columns", [])]
        rows = []
        for row in grid_info.get("Rows", []):
            values = row.get("Values", {})
            # Remove internal markers
            clean = {k: v for k, v in values.items() if not k.startswith("<")}
            if any(v for v in clean.values()):  # skip blank rows
                rows.append(clean)
        data["grids"][grid_name] = {"columns": columns, "rows": rows}

    # Extract top-level fields
    _extract_fields(form, data["fields"])

    # Extract tab fields
    for tab_name, tab_info in form.get("Tab", {}).items():
        if not isinstance(tab_info, dict):
            continue
        tab_data = {
            "text": tab_info.get("Text", tab_name),
            "fields": {},
            "grids": {},
        }
        children = tab_info.get("Children", {})
        _extract_fields(children, tab_data["fields"])
        # Sub-grids inside tabs
        for sg_name, sg_info in children.get("Grid", {}).items():
            if isinstance(sg_info, dict) and "Columns" in sg_info:
                cols = [c.get("Label", c.get("Name", "")) for c in sg_info.get("Columns", [])]
                sg_rows = []
                for r in sg_info.get("Rows", []):
                    vals = {k: v for k, v in r.get("Values", {}).items() if not k.startswith("<")}
                    if any(v for v in vals.values()):
                        sg_rows.append(vals)
                tab_data["grids"][sg_name] = {"columns": cols, "rows": sg_rows}
        data["tabs"][tab_name] = tab_data

    return data


def _collect_all_grid_rows(client, form_result):
    """
    Collect ALL grid rows by paginating with form_click_control(LoadNextPage).

    The MCP form tools return max 25 rows per page.  The grid response
    includes Pagination.HasNextPage and ActionIds: ["LoadNextPage"].
    We click LoadNextPage repeatedly until HasNextPage is False.

    Returns: (grid_name, all_rows, form_section_from_last_page)
    """
    form = form_result.get("FormState", {}).get("Form", {})

    # Find the main grid
    grid_name = None
    for gn, gi in form.get("Grid", {}).items():
        if isinstance(gi, dict) and "Rows" in gi:
            grid_name = gn
            break

    if not grid_name:
        return None, [], form

    def _extract_page_rows(frm):
        gi = frm.get("Grid", {}).get(grid_name, {})
        if not isinstance(gi, dict):
            return [], {}
        raw_rows = gi.get("Rows", [])
        pag = gi.get("Pagination", {})
        clean_rows = []
        for row in raw_rows:
            vals = row.get("Values", {})
            clean = {k: v for k, v in vals.items() if not k.startswith("<")}
            if any(v for v in clean.values()):
                clean_rows.append(clean)
        return clean_rows, pag

    # Page 1
    page_rows, pagination = _extract_page_rows(form)
    all_rows = list(page_rows)
    page_num = 1

    # Paginate
    while str(pagination.get("HasNextPage", "False")).lower() == "true":
        try:
            next_result = mcp_call(client, "form_click_control", {
                "controlName": grid_name,
                "actionId": "LoadNextPage",
            })
            if "raw" in next_result:
                print(f"      [pagination] page {page_num+1} returned raw — stopping")
                break  # pagination failed, keep what we have
            page_num += 1
            form = next_result.get("FormState", {}).get("Form", {})
            page_rows, pagination = _extract_page_rows(form)
            if not page_rows:
                print(f"      [pagination] page {page_num} empty — stopping")
                break
            all_rows.extend(page_rows)
        except Exception as e:
            print(f"      [pagination] page {page_num+1} error: {e}")
            break  # stop on error, keep collected rows

    if page_num > 1:
        print(f"      [pagination] {page_num} pages, {len(all_rows)} total rows")
    elif len(all_rows) == 25:
        # Suspicious: exactly 25 rows but HasNextPage was not True — try anyway
        hnp = pagination.get("HasNextPage", "MISSING")
        print(f"      [pagination] 25 rows but HasNextPage={hnp} — trying LoadNextPage anyway")
        try:
            next_result = mcp_call(client, "form_click_control", {
                "controlName": grid_name,
                "actionId": "LoadNextPage",
            })
            if "raw" not in next_result:
                form = next_result.get("FormState", {}).get("Form", {})
                extra_rows, pagination = _extract_page_rows(form)
                if extra_rows:
                    all_rows.extend(extra_rows)
                    page_num = 2
                    while str(pagination.get("HasNextPage", "False")).lower() == "true":
                        try:
                            next_result = mcp_call(client, "form_click_control", {
                                "controlName": grid_name,
                                "actionId": "LoadNextPage",
                            })
                            if "raw" in next_result:
                                break
                            page_num += 1
                            form = next_result.get("FormState", {}).get("Form", {})
                            extra_rows, pagination = _extract_page_rows(form)
                            if not extra_rows:
                                break
                            all_rows.extend(extra_rows)
                        except Exception:
                            break
                    print(f"      [pagination] recovered! {page_num} pages, {len(all_rows)} total rows")
        except Exception:
            print(f"      [pagination] LoadNextPage attempt failed — keeping 25 rows")

    return grid_name, all_rows, form


def read_form_records(client, form_result, include_details=True, tab_hint=None):
    """
    Read all records from a form.

    1. Gets grid rows from the open response (paginates to get ALL rows)
    2. If include_details, re-opens pagination from page 1 and selects
       each row to read Input/Checkbox/Tab values
    3. Returns list of flat dicts (one per record)

    If tab_hint is provided (e.g. "Transportation rate assignments"),
    looks for a tab whose text fuzzy-matches the hint and reads from
    that tab's sub-grid instead of the main grid.

    Returns: (records_list, column_names_list)
    """
    form_data = extract_form_data(client, form_result)

    # ── If tab_hint given, try to find matching tab sub-grid first ──
    if tab_hint:
        hint_lower = tab_hint.lower().strip()
        best_tab = None
        best_tab_grid = None
        best_score = 0

        for tab_name, tab_info in form_data["tabs"].items():
            tab_text = tab_info.get("text", "").lower().strip()
            # Check if hint matches tab text
            score = 0
            if tab_text == hint_lower:
                score = 100
            elif hint_lower in tab_text or tab_text in hint_lower:
                score = 60
            else:
                # Check word overlap
                hint_words = set(hint_lower.split())
                tab_words = set(tab_text.split())
                overlap = hint_words & tab_words
                if len(overlap) >= 2:
                    score = 30 + len(overlap) * 10

            if score > best_score and tab_info.get("grids"):
                best_score = score
                best_tab = tab_name
                # Pick the first sub-grid with data (or any sub-grid)
                for sg_name, sg_info in tab_info["grids"].items():
                    best_tab_grid = (sg_name, sg_info)
                    break

        if best_tab and best_tab_grid:
            sg_name, sg_info = best_tab_grid
            rows = sg_info.get("rows", [])

            # Paginate the tab sub-grid if it has 25 rows (might have more pages)
            if len(rows) >= 25:
                all_rows = list(rows)
                page_num = 1
                # Check if pagination is available on this sub-grid
                raw_form = form_result.get("FormState", {}).get("Form", {})
                # Try to find pagination info from the tab's grid
                while True:
                    try:
                        next_result = mcp_call(client, "form_click_control", {
                            "controlName": sg_name,
                            "actionId": "LoadNextPage",
                        })
                        if "raw" in next_result:
                            break
                        page_num += 1
                        next_form = next_result.get("FormState", {}).get("Form", {})
                        # Extract rows from the sub-grid in the new page
                        page_rows = []
                        for tab_name_inner, tab_info_inner in next_form.get("Tab", {}).items():
                            if not isinstance(tab_info_inner, dict):
                                continue
                            children = tab_info_inner.get("Children", {})
                            for sg_inner, sg_data in children.get("Grid", {}).items():
                                if sg_inner == sg_name and isinstance(sg_data, dict):
                                    for row in sg_data.get("Rows", []):
                                        vals = {k: v for k, v in row.get("Values", {}).items() if not k.startswith("<")}
                                        if any(v for v in vals.values()):
                                            page_rows.append(vals)
                        if not page_rows:
                            break
                        all_rows.extend(page_rows)
                        # Check HasNextPage
                        has_next = False
                        for tab_name_inner, tab_info_inner in next_form.get("Tab", {}).items():
                            if not isinstance(tab_info_inner, dict):
                                continue
                            children = tab_info_inner.get("Children", {})
                            for sg_inner, sg_data in children.get("Grid", {}).items():
                                if sg_inner == sg_name and isinstance(sg_data, dict):
                                    pag = sg_data.get("Pagination", {})
                                    has_next = str(pag.get("HasNextPage", "False")).lower() == "true"
                        if not has_next:
                            break
                    except Exception:
                        break
                if page_num > 1:
                    print(f"      [tab-pagination] {page_num} pages, {len(all_rows)} total rows")
                rows = all_rows

            cols = set()
            for r in rows:
                cols.update(r.keys())
            tab_text = form_data["tabs"][best_tab]["text"]
            print(f"      [tab-match] hint='{tab_hint}' -> tab='{tab_text}' grid={sg_name} ({len(rows)} rows)")
            return rows, sorted(cols)

    # Find the main grid — check top-level first, then sub-grids in tabs
    if not form_data["grids"]:
        # Check for sub-grids inside tabs
        tab_grid_name = None
        tab_grid_rows = []
        for tab_name, tab_info in form_data["tabs"].items():
            for sg_name, sg_info in tab_info.get("grids", {}).items():
                if sg_info.get("rows"):
                    tab_grid_name = sg_name
                    tab_grid_rows = sg_info["rows"]
                    break
            if tab_grid_name:
                break

        if tab_grid_rows:
            # Found data in a tab sub-grid — paginate if 25+ rows
            if len(tab_grid_rows) >= 25:
                all_rows = list(tab_grid_rows)
                page_num = 1
                while True:
                    try:
                        next_result = mcp_call(client, "form_click_control", {
                            "controlName": tab_grid_name,
                            "actionId": "LoadNextPage",
                        })
                        if "raw" in next_result:
                            break
                        page_num += 1
                        next_form = next_result.get("FormState", {}).get("Form", {})
                        page_rows = []
                        for tn, ti in next_form.get("Tab", {}).items():
                            if not isinstance(ti, dict):
                                continue
                            children = ti.get("Children", {})
                            for sg_n, sg_d in children.get("Grid", {}).items():
                                if sg_n == tab_grid_name and isinstance(sg_d, dict):
                                    for row in sg_d.get("Rows", []):
                                        vals = {k: v for k, v in row.get("Values", {}).items() if not k.startswith("<")}
                                        if any(v for v in vals.values()):
                                            page_rows.append(vals)
                        if not page_rows:
                            break
                        all_rows.extend(page_rows)
                        # Check HasNextPage
                        has_next = False
                        for tn, ti in next_form.get("Tab", {}).items():
                            if not isinstance(ti, dict):
                                continue
                            children = ti.get("Children", {})
                            for sg_n, sg_d in children.get("Grid", {}).items():
                                if sg_n == tab_grid_name and isinstance(sg_d, dict):
                                    pag = sg_d.get("Pagination", {})
                                    has_next = str(pag.get("HasNextPage", "False")).lower() == "true"
                        if not has_next:
                            break
                    except Exception:
                        break
                if page_num > 1:
                    print(f"      [tab-fallback-pagination] {page_num} pages, {len(all_rows)} total rows")
                tab_grid_rows = all_rows

            cols = set()
            for r in tab_grid_rows:
                cols.update(r.keys())
            return tab_grid_rows, sorted(cols)

        # No grid anywhere — this is a parameter/settings form (single record)
        record = {}
        for fname, finfo in form_data["fields"].items():
            label = finfo.get("label", fname)
            record[label] = finfo.get("value", "")
        for tab_name, tab_info in form_data["tabs"].items():
            for fname, finfo in tab_info["fields"].items():
                label = finfo.get("label", fname)
                record[label] = finfo.get("value", "")
        if record:
            return [record], list(record.keys())
        return [], []

    # Collect ALL rows across pages (this paginates to the last page)
    grid_name, grid_rows, _ = _collect_all_grid_rows(client, form_result)

    if not grid_name or not grid_rows:
        # Fallback to first page only
        grid_name = list(form_data["grids"].keys())[0]
        grid = form_data["grids"][grid_name]
        grid_rows = grid["rows"]

    columns = set()
    if grid_rows:
        for r in grid_rows:
            columns.update(r.keys())

    if not grid_rows:
        return [], list(columns)

    if not include_details:
        return grid_rows, list(grid_rows[0].keys()) if grid_rows else []

    # ── Detail read: select each row to get Input/Checkbox/Tab values ──
    # After pagination, the grid cursor is on the last page.
    # We need to navigate back to page 1 and walk through all rows.
    # Strategy: use form_select_grid_row — the MCP server handles
    # navigation internally. But if row index > current page size,
    # it returns "less than" error. So we paginate manually and
    # select rows page-by-page.
    records = []
    total_rows = len(grid_rows)
    page_size = 25  # MCP default page size

    # If total rows <= page_size, the grid is still on page 1 — no issue.
    # If total rows > page_size, we already paginated past page 1.
    # Either way, we need to re-open from row 1 by going back to first page.
    if total_rows > page_size:
        # Go back to first page by clicking FirstPage if available,
        # otherwise we just read grid-only data (no detail fields)
        try:
            first_page = mcp_call(client, "form_click_control", {
                "controlName": grid_name,
                "actionId": "LoadFirstPage",
            })
            if "raw" in first_page:
                # LoadFirstPage not available — return grid-only data
                print(f"      [detail] LoadFirstPage unavailable, using grid-only for {total_rows} rows")
                return grid_rows, sorted(columns)
        except Exception:
            print(f"      [detail] Cannot reset to page 1, using grid-only for {total_rows} rows")
            return grid_rows, sorted(columns)

    # Now walk through rows page by page
    current_page_start = 0  # 0-based index into grid_rows

    while current_page_start < total_rows:
        page_end = min(current_page_start + page_size, total_rows)

        for row_idx_in_page in range(1, page_end - current_page_start + 1):
            global_idx = current_page_start + row_idx_in_page - 1
            try:
                sel_result = mcp_call(client, "form_select_grid_row", {
                    "gridName": grid_name,
                    "rowNumber": row_idx_in_page,
                    "marking": "Unmarked",
                })

                if "raw" in sel_result:
                    raw = sel_result.get("raw", "")
                    if "less than" in raw:
                        break
                    records.append(dict(grid_rows[global_idx]))
                    continue

                # Start with grid row values
                record = dict(grid_rows[global_idx])

                # Add detail fields from Input/Checkbox
                sel_form = sel_result.get("FormState", {}).get("Form", {})
                detail_fields = {}
                _extract_fields(sel_form, detail_fields)

                for fname, finfo in detail_fields.items():
                    label = finfo.get("label", fname)
                    val = finfo.get("value", "")
                    if label not in record:
                        record[label] = val
                        columns.add(label)

                # Add tab fields
                for tab_name, tab_info in sel_form.get("Tab", {}).items():
                    if not isinstance(tab_info, dict):
                        continue
                    children = tab_info.get("Children", {})
                    tab_fields = {}
                    _extract_fields(children, tab_fields)
                    for fname, finfo in tab_fields.items():
                        label = finfo.get("label", fname)
                        val = finfo.get("value", "")
                        if label not in record:
                            record[label] = val
                            columns.add(label)

                records.append(record)
                time.sleep(0.15)

            except Exception as e:
                if "less than" in str(e):
                    break
                records.append(dict(grid_rows[global_idx]))

        # Move to next page if there are more rows
        current_page_start = page_end
        if current_page_start < total_rows:
            try:
                next_result = mcp_call(client, "form_click_control", {
                    "controlName": grid_name,
                    "actionId": "LoadNextPage",
                })
                if "raw" in next_result:
                    # Can't paginate further — append remaining as grid-only
                    for idx in range(current_page_start, total_rows):
                        records.append(dict(grid_rows[idx]))
                    break
            except Exception:
                for idx in range(current_page_start, total_rows):
                    records.append(dict(grid_rows[idx]))
                break

    return records, sorted(columns)


def _extract_fields(form_section, target_dict):
    """Extract Input, Checkbox, Combobox fields into target_dict."""
    # Inputs
    for fname, finfo in form_section.get("Input", {}).items():
        if isinstance(finfo, dict):
            target_dict[fname] = {
                "label": finfo.get("Label", fname),
                "value": finfo.get("Value", ""),
                "type": "input",
            }
    # Checkboxes
    for fname, finfo in form_section.get("Checkbox", {}).items():
        if isinstance(finfo, dict):
            target_dict[fname] = {
                "label": finfo.get("Label", fname),
                "value": finfo.get("IsChecked", ""),
                "type": "checkbox",
            }
    # Comboboxes
    for fname, finfo in form_section.get("Combobox", {}).items():
        if isinstance(finfo, dict):
            target_dict[fname] = {
                "label": finfo.get("Label", fname),
                "value": finfo.get("Value", finfo.get("SelectedValue", "")),
                "type": "combobox",
            }


def _clean_title(title):
    """
    Clean a config deliverable title into path segments and a search-ready
    last segment. Returns (segments, last_segment, paren_hint).
    """
    clean = title.strip()
    # 1. Strip stacked leading bracket prefixes: [CRO], [FDD-FIN-ED-02], [CUTOVER], etc.
    while clean.startswith("["):
        clean = re.sub(r'^\[.*?\]\s*', '', clean, count=1)
    # 2. Strip "INT:" or similar uppercase prefix
    clean = re.sub(r'^[A-Z]+:\s*', '', clean)
    # 3. Normalise "->" to ">"
    clean = clean.replace('->', '>')

    if '>' not in clean:
        return [], clean.strip(), None

    parts = [p.strip() for p in clean.split('>')]
    parts = [p for p in parts if p]  # drop empties

    last = parts[-1] if parts else ""

    # 4. Handle bracket-only segments like "[Detour Manual Setup]"
    #    If the entire last segment is a bracket token, extract its content
    #    as the search term instead of stripping it to empty.
    bracket_only_match = re.fullmatch(r'\[([^\]]+)\]', last.strip())
    if bracket_only_match:
        last = bracket_only_match.group(1).strip()
    else:
        # Strip trailing bracket tokens: [HR01], [Empties], [Export], [Import], [Global]
        last = re.sub(r'\s*\[.*?\]', '', last).strip()

    # 5. Strip " - free text description" suffix
    if ' - ' in last:
        last = last.split(' - ')[0].strip()
    # 6. Extract parenthetical hint
    paren_match = re.search(r'\(([^)]+)\)', last)
    paren_hint = paren_match.group(1).strip() if paren_match else None
    # 7. Strip parenthetical from the main segment
    base_last = re.sub(r'\s*\([^)]*\)', '', last).strip()

    # 8. If base_last is empty after cleaning, fall back to parent segment
    if not base_last and len(parts) >= 2:
        base_last = re.sub(r'\s*\[.*?\]', '', parts[-2]).strip()
        print(f"      [clean_title] last segment was empty, using parent: '{base_last}'")

    # Update the parts list with cleaned last segment
    if parts:
        parts[-1] = base_last

    return parts, base_last, paren_hint


# Words that are too generic to search alone — need parent context
_GENERIC_TERMS = {
    "parameters", "setup", "posting", "operations", "configuration",
    "settings", "number sequences", "categories", "terms of payment",
    "reasons", "types", "codes", "groups", "methods", "profiles",
    "journal names", "dimensions", "calendars", "policies",
    "sequences", "rules", "intervals", "charges", "schedules",
    "aging periods", "port", "templates", "reconciliation reasons",
}


def _score_result(item, parts, base_last, paren_hint):
    """
    Score a search result against the full navigation path context.
    Higher score = better match. Returns int score.
    """
    name_lower = item["name"].lower()
    text_lower = item["text"].lower()
    score = 0

    # Exact text match on the base last segment
    if text_lower == base_last.lower():
        score += 100
    elif base_last.lower() in text_lower:
        score += 50

    # Parenthetical hint match (high value — very specific)
    if paren_hint:
        if text_lower == paren_hint.lower():
            score += 120
        elif paren_hint.lower() in text_lower:
            score += 60

    # Path context scoring — check if parent module keywords appear in the
    # menu item name/text (e.g., "ProdParameters" for "Cost management > ... > Parameters")
    module_keywords = _extract_module_keywords(parts)
    for kw in module_keywords:
        kw_l = kw.lower()
        if kw_l in name_lower:
            score += 30
        if kw_l in text_lower:
            score += 15

    # Penalise very short names (usually wrong generic matches)
    if len(item["name"]) <= 3:
        score -= 20

    return score


def _extract_module_keywords(parts):
    """
    Extract disambiguating keywords from the navigation path segments.
    E.g., ["Cost management", "Manufacturing policies setup", "Parameters"]
    → ["cost", "manufacturing", "production", "prod"]
    """
    # Module name mappings — D365 module prefixes that appear in menu item names
    _MODULE_HINTS = {
        "cost management": ["cost", "prod", "production"],
        "cost accounting": ["costaccounting", "cost"],
        "general ledger": ["ledger", "gl"],
        "accounts receivable": ["cust", "customer", "ar"],
        "accounts payable": ["vend", "vendor", "ap"],
        "inventory management": ["invent", "inventory"],
        "warehouse management": ["whs", "warehouse"],
        "procurement": ["purch", "procurement"],
        "sales": ["sales"],
        "project management": ["proj", "project"],
        "production control": ["prod", "production"],
        "transportation management": ["tms", "transport"],
        "tax": ["tax", "salestax"],
        "fixed assets": ["asset", "fixedasset"],
        "credit and collections": ["credit", "collection"],
        "cash and bank management": ["bank", "cash"],
        "budgeting": ["budget"],
        "organization administration": ["org", "admin"],
        "system administration": ["sys", "system"],
        "human resources": ["hcm", "hr"],
        "retail": ["retail"],
        "master planning": ["req", "masterplan", "planning"],
        "product information management": ["product", "ecoresproduct"],
        "pricing management": ["pricing", "price"],
        "globalization studio": ["globalization", "er"],
        "asset management": ["entasset", "asset"],
        "data management": ["dmf", "data", "framework"],
        "fleet management": ["fleet"],
        "subscription billing": ["subscription", "billing"],
        "revenue recognition": ["revenue"],
        "lease accounting": ["lease"],
        "rebate management": ["rebate"],
    }

    keywords = []
    for part in parts[:-1]:  # skip last segment
        part_lower = part.lower().strip()
        # Direct module hint lookup
        for mod_key, hints in _MODULE_HINTS.items():
            if mod_key in part_lower:
                keywords.extend(hints)
                break
        # Also add significant words from the path segments
        for word in part_lower.split():
            if len(word) > 4 and word not in ("setup", "periodic", "tasks", "inquiries", "reports"):
                keywords.append(word)

    return list(set(keywords))


def resolve_menu_item_from_path(client, title, company_id="HR01"):
    """
    Resolve a D365 navigation path (from config deliverable title)
    to an actual menu item name using multi-strategy search.

    Returns: (menu_item_name, menu_item_type, paren_hint) or (None, None, None)
    """
    parts, base_last, paren_hint = _clean_title(title)

    # No navigation path → free-text title, can't resolve
    if not parts or len(parts) < 2:
        return None, None, None

    # Build ordered list of search strategies
    searches = []

    # Strategy 1: parenthetical hint (very specific, e.g. "Freight bill types")
    if paren_hint and len(paren_hint) >= 3:
        searches.append(paren_hint)

    # Strategy 2: base last segment (e.g. "Sales tax groups")
    if base_last and len(base_last) >= 3:
        searches.append(base_last)

    # Strategy 3: for generic terms, combine with parent segment
    if base_last.lower().strip() in _GENERIC_TERMS and len(parts) >= 2:
        parent = re.sub(r'\s*\[.*?\]', '', parts[-2]).strip()
        combined = f"{parent} {base_last}"
        searches.append(combined)

    # Strategy 4: second-to-last segment alone (when last is very generic)
    if base_last.lower().strip() in _GENERIC_TERMS and len(parts) >= 2:
        parent = re.sub(r'\s*\[.*?\]', '', parts[-2]).strip()
        if len(parent) >= 3:
            searches.append(parent)

    # Strategy 5: individual significant words from the last segment
    #   e.g. "Compatibility options" → try "Compatibility"
    #   e.g. "Job history cleanup" → try "cleanup", "history"
    if base_last and len(base_last) >= 3:
        words = base_last.split()
        for word in words:
            if len(word) >= 5 and word.lower() not in ("setup", "management", "information", "configuration"):
                searches.append(word)

    # Strategy 6: parent segment as search term (even when last is not generic)
    #   e.g. "Framework parameters" when "Compatibility options" finds nothing
    if len(parts) >= 3:
        parent = re.sub(r'\s*\[.*?\]', '', parts[-2]).strip()
        if parent and len(parent) >= 3 and parent.lower().strip() not in _GENERIC_TERMS:
            searches.append(parent)

    # Deduplicate while preserving order
    seen = set()
    unique_searches = []
    for s in searches:
        s_key = s.lower().strip()
        if s_key not in seen:
            seen.add(s_key)
            unique_searches.append(s)

    # Collect ALL candidates from all search strategies, score them
    all_candidates = []
    searched = set()

    for search in unique_searches:
        search_key = search.lower().strip()
        if search_key in searched:
            continue
        searched.add(search_key)

        try:
            items = find_menu_item(client, search, company_id)
            if items:
                for item in items:
                    # Avoid duplicate candidates
                    item_key = (item["name"], item["type"])
                    if item_key not in {(c["name"], c["type"]) for c in all_candidates}:
                        all_candidates.append(item)
        except Exception:
            continue

    if not all_candidates:
        print(f"      [resolve] NO candidates for '{title}' (searched: {list(searched)})")
        return None, None, None

    # Score all candidates against the full path context
    scored = []
    for item in all_candidates:
        s = _score_result(item, parts, base_last, paren_hint)
        scored.append((s, item))

    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best = scored[0]

    # Require minimum score to avoid completely wrong matches
    if best_score < 5:
        top3 = [(s, i["name"], i["text"]) for s, i in scored[:3]]
        print(f"      [resolve] Low scores for '{title}': {top3}")
        return None, None, None

    return best["name"], best["type"], paren_hint
