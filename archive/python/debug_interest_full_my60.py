"""
Comprehensive dump of ALL controls with values for Interest form on MY60.
Includes grid row data, section grouping, and every control that has any value.
"""
import sys
import json
import time

sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient
from form_reader import open_form, close_form

with open(r"C:\D365DataValidator\config.json") as f:
    config = json.load(f)

SEARCH_TERMS = list("abcdefghijklmnopqrstuvwxyz") + [
    "0", "1", "2", "3", "4", "5", "6", "7", "8", "9",
    "Code", "Name", "Date", "Num", "Type", "Status",
    "Amount", "Qty", "Unit", "Group", "Account", "Description",
    "Currency", "Company", "Ledger", "Dimension", "Table", "Line",
    "Cust", "Interest", "Fee", "Calendar", "Paym", "Rate", "Percent",
    "Invoice", "Calculate", "Grace", "Period", "Range", "Method",
    "Earnings", "Payments", "Note", "Voucher", "Post", "Charge",
    "Min", "Max", "Day", "Month", "From", "To", "Base", "Value",
    "Debit", "Credit", "Text", "Language", "Version", "Valid",
    "Single", "Effective", "Expiration", "Delete", "New", "Save",
    "Business", "Tab", "Send", "Event", "Filter", "Action",
    "Tax", "Item", "Sales", "Purchase", "Remove", "Close",
]

client = D365McpClient(config["environments"]["ENV1"])
client.connect()

print("Opening Interest form on MY60...")
form_result = open_form(client, "Interest", "Display", "MY60")
fs = form_result.get("FormState", {})
print(f"Form: {fs.get('Name')} — {fs.get('Caption')}")
print(f"Company: {fs.get('Company')}")

# ── 1. Dump full initial FormState structure ──
form_obj = fs.get("Form", {})

print("\n" + "="*80)
print("SECTION A: FULL FORM STRUCTURE (from FormState)")
print("="*80)

def dump_form_structure(obj, indent=0, path="Form"):
    """Recursively dump form structure showing all controls and their values."""
    if not isinstance(obj, dict):
        return
    for key, val in obj.items():
        prefix = "  " * indent
        if key == "Grid" and isinstance(val, dict):
            for gname, ginfo in val.items():
                if not isinstance(ginfo, dict):
                    continue
                print(f"\n{prefix}[GRID] {gname}")
                print(f"{prefix}  Text: {ginfo.get('Text', '')}")
                cols = ginfo.get("Columns", [])
                print(f"{prefix}  Columns ({len(cols)}):")
                for col in cols:
                    print(f"{prefix}    - {col.get('Name', '')} (Label: {col.get('Label', '')})")
                rows = ginfo.get("Rows", [])
                print(f"{prefix}  Rows ({len(rows)}):")
                for row in rows:
                    rn = row.get("RowNumber", "?")
                    vals = row.get("Values", {})
                    print(f"{prefix}    Row {rn}: {vals}")
                pag = ginfo.get("Pagination", {})
                if pag:
                    print(f"{prefix}  Pagination: HasNext={pag.get('HasNextPage')}, HasPrev={pag.get('HasPrevPage')}")

        elif key in ("Input", "Combobox", "Checkbox", "RealInput", "IntInput",
                      "DateInput", "TimeInput", "SegmentedEntry"):
            if isinstance(val, dict):
                for cname, cinfo in val.items():
                    if isinstance(cinfo, dict):
                        label = cinfo.get("Label", "")
                        value = cinfo.get("Value", "")
                        vtext = cinfo.get("ValueText", "")
                        display = vtext if vtext else value
                        req = cinfo.get("IsRequired", "")
                        edit = cinfo.get("IsEditable", "")
                        has_lookup = cinfo.get("HasLookup", "")
                        options = cinfo.get("Options", [])
                        marker = " *** HAS VALUE ***" if display and str(display).strip() else ""
                        print(f"{prefix}[{key}] {cname}{marker}")
                        print(f"{prefix}  Label: {label}")
                        print(f"{prefix}  Value: {display}")
                        if req: print(f"{prefix}  Required: {req}")
                        if edit: print(f"{prefix}  Editable: {edit}")
                        if has_lookup: print(f"{prefix}  HasLookup: {has_lookup}")
                        if options:
                            opts_str = ", ".join(f"{o.get('Label')}={o.get('Value')}" for o in options)
                            print(f"{prefix}  Options: {opts_str}")

        elif key == "Tab" and isinstance(val, dict):
            for tname, tinfo in val.items():
                if isinstance(tinfo, dict):
                    tlabel = tinfo.get("Label", "")
                    print(f"\n{prefix}[TAB] {tname} — {tlabel}")
                    dump_form_structure(tinfo, indent + 1, f"{path}.Tab.{tname}")

        elif key == "TabPage" and isinstance(val, dict):
            for tname, tinfo in val.items():
                if isinstance(tinfo, dict):
                    tlabel = tinfo.get("Label", "")
                    print(f"\n{prefix}[TABPAGE] {tname} — {tlabel}")
                    dump_form_structure(tinfo, indent + 1, f"{path}.TabPage.{tname}")

        elif key == "Group" and isinstance(val, dict):
            for gname, ginfo in val.items():
                if isinstance(ginfo, dict):
                    glabel = ginfo.get("Label", "")
                    print(f"\n{prefix}[GROUP] {gname} — {glabel}")
                    dump_form_structure(ginfo, indent + 1, f"{path}.Group.{gname}")

        elif key == "Children" and isinstance(val, dict):
            dump_form_structure(val, indent, f"{path}.Children")

        elif key == "Button" and isinstance(val, dict):
            for bname, binfo in val.items():
                if isinstance(binfo, dict):
                    blabel = binfo.get("Label", "")
                    print(f"{prefix}[BUTTON] {bname} — {blabel}")

        elif isinstance(val, dict) and key not in (
            "Rows", "Cells", "Pagination", "MenuButton", "ButtonGroup",
            "Label", "HelpText", "Value", "ValueText", "IsRequired",
            "IsEditable", "HasLookup", "Options", "Columns", "Text",
            "Name", "Guidance", "RequiresLookup",
        ):
            # Unknown section — show it
            if any(isinstance(v, dict) for v in val.values()):
                print(f"\n{prefix}[SECTION] {key}")
                dump_form_structure(val, indent + 1, f"{path}.{key}")

dump_form_structure(form_obj)

# ── 2. Select first row and dump expanded state ──
print("\n\n" + "="*80)
print("SECTION B: AFTER SELECTING FIRST GRID ROW")
print("="*80)

try:
    sel_result = client.call_tool("form_select_grid_row", {
        "gridName": "Interest", "rowNumber": "1", "marking": "Unmarked",
    })
    if isinstance(sel_result, str) and sel_result.strip():
        sel_parsed = json.loads(sel_result)
        if isinstance(sel_parsed, dict):
            sel_form = sel_parsed.get("FormState", {}).get("Form", {})
            if sel_form:
                print("\nExpanded form after row selection:")
                dump_form_structure(sel_form)
            else:
                print("No expanded Form in response")
        elif isinstance(sel_parsed, list):
            print(f"Got list response with {len(sel_parsed)} items")
    time.sleep(0.5)
except Exception as e:
    print(f"Row select failed: {e}")

# ── 3. form_find_controls — ALL controls ──
print("\n\n" + "="*80)
print("SECTION C: ALL CONTROLS VIA form_find_controls")
print("="*80)

all_controls = {}
for term in SEARCH_TERMS:
    try:
        raw = client.call_tool("form_find_controls", {"controlSearchTerm": term})
        if isinstance(raw, str) and raw.strip():
            parsed = json.loads(raw)
            if isinstance(parsed, list):
                for item in parsed:
                    cname = item.get("Name", "")
                    if cname and cname not in all_controls:
                        all_controls[cname] = item
    except Exception:
        pass

print(f"\nTotal unique controls from form_find_controls: {len(all_controls)}")

# Print ALL controls grouped by type
grids = {}
fields_with_val = {}
fields_no_val = {}
buttons = {}

for cname, item in sorted(all_controls.items()):
    props = item.get("Properties", {})
    if not isinstance(props, dict):
        continue

    # Check if it's a grid
    if "Columns" in props:
        grids[cname] = props
        continue

    label = props.get("Label", "")
    value = props.get("Value", "")
    value_text = props.get("ValueText", "")
    display = value_text if value_text else value

    # Is it a button?
    if "Button" in cname or (not label and not value and props.get("HelpText")):
        buttons[cname] = props
        continue

    if display and str(display).strip():
        fields_with_val[cname] = {
            "label": label, "value": str(display).strip(),
            "raw_value": value, "is_required": props.get("IsRequired"),
            "is_editable": props.get("IsEditable"),
            "has_lookup": props.get("HasLookup"),
            "options": props.get("Options", []),
        }
    else:
        fields_no_val[cname] = {
            "label": label,
            "is_required": props.get("IsRequired"),
            "is_editable": props.get("IsEditable"),
            "has_lookup": props.get("HasLookup"),
        }

print(f"\n--- GRIDS ({len(grids)}) ---")
for gname, gprops in grids.items():
    cols = gprops.get("Columns", [])
    rows = gprops.get("Rows", [])
    print(f"\n  [{gname}] — {gprops.get('Text', '')}")
    print(f"    Columns: {len(cols)}")
    for col in cols:
        print(f"      {col.get('Name', '')} — {col.get('Label', '')}")
    print(f"    Rows: {len(rows)}")
    for row in rows:
        print(f"      Row {row.get('RowNumber', '?')}: {row.get('Values', {})}")
    pag = gprops.get("Pagination", {})
    if pag:
        print(f"    Pagination: HasNext={pag.get('HasNextPage')}, HasPrev={pag.get('HasPrevPage')}")

print(f"\n--- FIELDS WITH VALUES ({len(fields_with_val)}) ---")
for cname, info in sorted(fields_with_val.items()):
    print(f"\n  {cname}")
    print(f"    Label:    {info['label']}")
    print(f"    Value:    {info['value']}")
    if info['raw_value'] != info['value']:
        print(f"    RawValue: {info['raw_value']}")
    if info['is_required']: print(f"    Required: {info['is_required']}")
    if info['is_editable']: print(f"    Editable: {info['is_editable']}")
    if info['has_lookup']: print(f"    Lookup:   {info['has_lookup']}")
    if info['options']:
        opts = ", ".join(f"{o.get('Label')}={o.get('Value')}" for o in info['options'])
        print(f"    Options:  {opts}")

print(f"\n--- FIELDS WITHOUT VALUES ({len(fields_no_val)}) ---")
for cname, info in sorted(fields_no_val.items()):
    print(f"  {cname} — Label: {info['label']}, Lookup: {info.get('has_lookup', '')}")

print(f"\n--- BUTTONS ({len(buttons)}) ---")
for bname, bprops in sorted(buttons.items()):
    print(f"  {bname} — {bprops.get('Label', '')}")

close_form(client)
