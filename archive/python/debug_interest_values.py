"""
Find all controls with values for Interest form on MY30.
"""
import sys
import json
import time

sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient
from form_reader import open_form, close_form

CONFIG_FILE = r"C:\D365DataValidator\config.json"

SEARCH_TERMS = list("abcdefghijklmnopqrstuvwxyz") + [
    "0", "1", "2", "Code", "Name", "Date", "Num", "Type", "Status",
    "Amount", "Qty", "Unit", "Group", "Account", "Description",
    "Currency", "Company", "Ledger", "Dimension", "Table", "Line",
    "Cust", "Interest", "Fee", "Calendar", "Paym", "Rate", "Percent",
    "Invoice", "Calculate", "Grace", "Period", "Range", "Method",
    "Earnings", "Payments", "Note", "Voucher", "Post", "Charge",
    "Min", "Max", "Day", "Month", "From", "To", "Base", "Value",
]


def main():
    with open(CONFIG_FILE) as f:
        config = json.load(f)

    client = D365McpClient(config["environments"]["ENV1"])
    client.connect()

    # Open Interest form
    print("Opening Interest form on MY30...")
    form_result = open_form(client, "Interest", "Display", "MY30")
    fs = form_result.get("FormState", {})
    print(f"Form: {fs.get('Name')} — {fs.get('Caption')}")

    # Select first row to expand details
    print("Selecting first grid row...")
    try:
        client.call_tool("form_select_grid_row", {
            "gridName": "Interest", "rowNumber": "1", "marking": "Unmarked",
        })
        time.sleep(0.5)
    except Exception as e:
        print(f"  Row select failed: {e}")

    # Open all tabs
    for tab in ["Earnings", "Payments", "General", "Setup", "Details",
                "Interest", "Note", "Ranges"]:
        try:
            client.call_tool("form_open_or_close_tab", {"tabName": tab})
        except Exception:
            pass

    # Collect all controls
    print("Searching all controls...")
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
                            all_controls[cname] = item.get("Properties", {})
        except Exception:
            pass

    # Also extract from initial form structure (grid columns)
    form_obj = fs.get("Form", {})
    _extract_from_form(form_obj, all_controls)

    print(f"\nTotal unique controls: {len(all_controls)}")

    # Find controls with values
    controls_with_values = []
    controls_without_values = []

    for cname, props in sorted(all_controls.items()):
        if not isinstance(props, dict):
            continue

        # Skip system/button controls
        if cname.startswith("SystemDefined") or "Button" in cname:
            continue

        # Check for value
        value = props.get("Value", "")
        value_text = props.get("ValueText", "")
        label = props.get("Label", "")

        # Grid controls have Rows/Columns instead of Value
        if "Columns" in props:
            row_count = len(props.get("Rows", []))
            controls_with_values.append({
                "name": cname,
                "label": label or props.get("Text", ""),
                "value": f"[Grid: {row_count} rows]",
                "type": "Grid",
            })
            continue

        display_val = value_text if value_text else value
        if display_val and str(display_val).strip():
            controls_with_values.append({
                "name": cname,
                "label": label,
                "value": str(display_val).strip(),
                "type": "Field",
            })
        else:
            controls_without_values.append({
                "name": cname,
                "label": label,
            })

    # Print results
    print(f"\n{'='*70}")
    print(f"CONTROLS WITH VALUES: {len(controls_with_values)}")
    print(f"{'='*70}")
    for c in controls_with_values:
        print(f"  [{c['type']:5}] {c['name']}")
        print(f"         Label: {c['label']}")
        print(f"         Value: {c['value']}")
        print()

    print(f"\n{'='*70}")
    print(f"CONTROLS WITHOUT VALUES: {len(controls_without_values)}")
    print(f"{'='*70}")
    for c in controls_without_values:
        print(f"  {c['name']} — {c['label']}")

    close_form(client)


def _extract_from_form(form_obj, controls, depth=0):
    if not isinstance(form_obj, dict) or depth > 8:
        return
    for key, val in form_obj.items():
        if not isinstance(val, dict):
            continue
        if key == "Grid":
            for gname, ginfo in val.items():
                if isinstance(ginfo, dict) and gname not in controls:
                    controls[gname] = ginfo
        elif key in ("Input", "Combobox", "Checkbox", "RealInput", "IntInput",
                      "DateInput", "TimeInput", "SegmentedEntry"):
            for cname, cinfo in val.items():
                if cname and cname not in controls:
                    controls[cname] = cinfo if isinstance(cinfo, dict) else {}
        elif key in ("Tab", "TabPage", "Children", "Group"):
            _extract_from_form(val, controls, depth + 1)
        elif key not in ("Rows", "Cells", "Button", "MenuButton",
                         "ButtonGroup", "Pagination"):
            _extract_from_form(val, controls, depth + 1)


if __name__ == "__main__":
    main()
