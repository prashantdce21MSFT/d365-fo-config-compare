"""Debug: open CustTable on MY30, select row 0, dump all tabs and fields."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient

with open(r"C:\D365DataValidator\config.json") as f:
    cfg = json.load(f)
client = D365McpClient(cfg["environments"]["ENV1"])
client.connect()

# Open form with legal entity
raw = client.call_tool("form_open_menu_item", {
    "menuItemName": "CustTable",
    "menuItemType": "Display",
    "legalEntity": "MY30"
})
parsed = json.loads(raw)
form = parsed.get("FormState", {}).get("Form", {})
print(f"Form opened. Caption: {parsed.get('FormState', {}).get('FormCaption', '')}")
print(f"Top-level form types: {sorted(form.keys())}")

# Check all Tab/TabPage at every level
def find_all_tabs(obj, path="Form", depth=0):
    if not isinstance(obj, dict) or depth > 6:
        return
    for ttype in ["Tab", "TabPage"]:
        for tname, tinfo in obj.get(ttype, {}).items():
            if isinstance(tinfo, dict):
                lbl = tinfo.get("Label", tinfo.get("Text", ""))
                print(f"  {'  '*depth}{path} > {ttype} > {tname}: label='{lbl}'")
                children = tinfo.get("Children", {})
                if isinstance(children, dict):
                    find_all_tabs(children, f"{tname}", depth+1)

print("\n=== ALL TABS IN INITIAL FORM ===")
find_all_tabs(form)

# Select first row
print("\n=== SELECTING ROW 0 ===")
try:
    sel_raw = client.call_tool("form_select_grid_row", {
        "gridName": "Grid", "rowNumber": "0", "marking": "Unmarked"
    })
    if sel_raw and sel_raw.strip():
        sel = json.loads(sel_raw)
        sel_form = sel.get("FormState", {}).get("Form", {})
        print(f"After row select, form types: {sorted(sel_form.keys())}")
        print("\n=== ALL TABS AFTER ROW SELECT ===")
        find_all_tabs(sel_form)
    else:
        print("Empty response from row select")
except Exception as e:
    print(f"Row select failed: {e}")

# Try opening various tab names that might be Acme Asia
print("\n=== TRYING TAB OPENS ===")
tab_candidates = [
    "CBATab", "TabCBA", "AcmeAsia", "TabAcmeAsia",
    "CBATabPage", "CHBTab", "TabCHB", "AcmeHuB", "TabAcmeHuB",
    "CHBTabPage", "CBACustomerTab", "CHBCustomerTab",
    "TabGeneral", "General", "TabHeader", "HeaderView",
    "TabLineDetails", "DetailsTab", "DetailView"
]
for tab_name in tab_candidates:
    try:
        tab_raw = client.call_tool("form_open_or_close_tab", {"tabName": tab_name, "tabAction": "Open"})
        if isinstance(tab_raw, str) and tab_raw.strip():
            tp = json.loads(tab_raw)
            fs = tp.get("FormState", {}).get("Form", {})
            result = tp.get("Result", "")
            if fs:
                print(f"  {tab_name}: OPENED! Finding tabs inside...")
                find_all_tabs(fs, "  ")
            elif "not found" in str(result).lower() or "error" in str(result).lower():
                pass  # skip not-found silently
            else:
                print(f"  {tab_name}: Result={str(result)[:80]}")
    except:
        pass

# Search for consignment control
print("\n=== SEARCHING FOR CONSIGNMENT ===")
for term in ["Consignment", "CHBConsignment", "CBAConsignment", "ConsignmentWarehouse"]:
    try:
        fc_raw = client.call_tool("form_find_controls", {"controlSearchTerm": term})
        if fc_raw and fc_raw.strip():
            controls = json.loads(fc_raw)
            if controls:
                for c in controls:
                    name = c.get("Name", "")
                    props = c.get("Properties", {})
                    val = props.get("Value", "")
                    lbl = props.get("Label", "")
                    print(f"  [{term}] {name}: label='{lbl}', value='{val}'")
    except:
        pass

client.call_tool("form_close_form", {})
print("\nDone.")
