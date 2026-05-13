"""Find all controls with values for Interest form on MY60."""
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
    "0", "1", "2", "Code", "Name", "Date", "Num", "Type", "Status",
    "Amount", "Qty", "Unit", "Group", "Account", "Description",
    "Currency", "Company", "Ledger", "Dimension", "Table", "Line",
    "Cust", "Interest", "Fee", "Calendar", "Paym", "Rate", "Percent",
    "Invoice", "Calculate", "Grace", "Period", "Range", "Method",
    "Earnings", "Payments", "Note", "Voucher", "Post", "Charge",
    "Min", "Max", "Day", "Month", "From", "To", "Base", "Value",
]

client = D365McpClient(config["environments"]["ENV1"])
client.connect()

print("Opening Interest form on MY60...")
form_result = open_form(client, "Interest", "Display", "MY60")
fs = form_result.get("FormState", {})
print(f"Form: {fs.get('Name')} — {fs.get('Caption')}")

# Select first row
try:
    client.call_tool("form_select_grid_row", {
        "gridName": "Interest", "rowNumber": "1", "marking": "Unmarked",
    })
    time.sleep(0.5)
except Exception as e:
    print(f"  Row select failed: {e}")

# Open tabs
for tab in ["Earnings", "Payments", "General", "Setup", "Details",
            "Interest", "Note", "Ranges"]:
    try:
        client.call_tool("form_open_or_close_tab", {"tabName": tab})
    except Exception:
        pass

# Collect controls
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

# Extract from form structure
def _extract(obj, ctrls, depth=0):
    if not isinstance(obj, dict) or depth > 8: return
    for k, v in obj.items():
        if not isinstance(v, dict): continue
        if k == "Grid":
            for gn, gi in v.items():
                if isinstance(gi, dict) and gn not in ctrls: ctrls[gn] = gi
        elif k in ("Input","Combobox","Checkbox","RealInput","IntInput","DateInput","TimeInput","SegmentedEntry"):
            for cn, ci in v.items():
                if cn and cn not in ctrls: ctrls[cn] = ci if isinstance(ci, dict) else {}
        elif k in ("Tab","TabPage","Children","Group"):
            _extract(v, ctrls, depth+1)
        elif k not in ("Rows","Cells","Button","MenuButton","ButtonGroup","Pagination"):
            _extract(v, ctrls, depth+1)

_extract(fs.get("Form", {}), all_controls)

# Analyze
with_values = []
without_values = []

for cname, props in sorted(all_controls.items()):
    if not isinstance(props, dict): continue
    if cname.startswith("SystemDefined") or "Button" in cname: continue

    value = props.get("Value", "")
    value_text = props.get("ValueText", "")
    label = props.get("Label", "")

    if "Columns" in props:
        row_count = len(props.get("Rows", []))
        with_values.append({"name": cname, "label": label or props.get("Text",""), "value": f"[Grid: {row_count} rows]", "type": "Grid"})
        continue

    display_val = value_text if value_text else value
    if display_val and str(display_val).strip():
        with_values.append({"name": cname, "label": label, "value": str(display_val).strip(), "type": "Field"})
    else:
        without_values.append({"name": cname, "label": label})

print(f"\nTotal controls: {len(all_controls)}")
print(f"\n{'='*70}")
print(f"CONTROLS WITH VALUES: {len(with_values)}")
print(f"{'='*70}")
for c in with_values:
    print(f"  [{c['type']:5}] {c['name']}")
    print(f"         Label: {c['label']}")
    print(f"         Value: {c['value']}")
    print()

print(f"\n{'='*70}")
print(f"CONTROLS WITHOUT VALUES: {len(without_values)}")
print(f"{'='*70}")
for c in without_values:
    print(f"  {c['name']} — {c['label']}")

close_form(client)
