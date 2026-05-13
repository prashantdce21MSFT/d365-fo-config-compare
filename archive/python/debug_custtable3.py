"""Debug: raw FormState inspection for CustTable after selecting row."""
import sys, io, json, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient
from form_reader import open_form, close_form, mcp_call

with open(r"C:\D365DataValidator\config.json") as f:
    cfg = json.load(f)
client = D365McpClient(cfg["environments"]["ENV1"])
client.connect()

# Open CustTable
print("Opening CustTable on MY30...")
form_result = open_form(client, "CustTable", "Display", "MY30")
print(f"Form result type: {type(form_result)}")

# Select row 1
print("\nSelecting row 1...")
sel_raw = client.call_tool("form_select_grid_row", {
    "gridName": "Grid",
    "rowNumber": "1",
    "marking": "Unmarked",
})

if isinstance(sel_raw, str) and sel_raw.strip():
    sel = json.loads(sel_raw)
    form = sel.get("FormState", {}).get("Form", {})

    # Dump ALL top-level keys in Form
    print(f"\nTop-level Form keys: {sorted(form.keys())}")

    # For each key, show its sub-keys
    def dump_structure(obj, prefix="", depth=0):
        if not isinstance(obj, dict) or depth > 4:
            return
        for key in sorted(obj.keys()):
            val = obj[key]
            if isinstance(val, dict):
                sub_keys = sorted(val.keys())
                if len(sub_keys) <= 10:
                    print(f"{'  '*depth}{prefix}{key}: dict with {len(sub_keys)} keys: {sub_keys}")
                else:
                    print(f"{'  '*depth}{prefix}{key}: dict with {len(sub_keys)} keys: {sub_keys[:10]}...")
                # If this looks like a control type (Tab, TabPage, Group, etc.), show children
                for sk in sub_keys[:5]:
                    sv = val[sk]
                    if isinstance(sv, dict):
                        inner_keys = sorted(sv.keys())
                        label = sv.get("Label", sv.get("Text", ""))
                        has_children = "Children" in sv
                        print(f"{'  '*(depth+1)}{sk}: label='{label}', children={has_children}, keys={inner_keys[:8]}")
            elif isinstance(val, list):
                print(f"{'  '*depth}{prefix}{key}: list with {len(val)} items")
            else:
                print(f"{'  '*depth}{prefix}{key}: {type(val).__name__} = {str(val)[:80]}")

    print("\n=== FORM STRUCTURE ===")
    dump_structure(form)

    # Also specifically search for anything that might be a FastTab
    # FastTabs in D365 are often "Tab" with TabStyle="FastTabs"
    print("\n=== LOOKING FOR FASTTAB INDICATORS ===")
    def search_for_tabs(obj, path="", depth=0):
        if not isinstance(obj, dict) or depth > 6:
            return
        for key in obj:
            val = obj[key]
            if key in ("Tab", "TabPage", "TabHeader", "FastTab", "Section", "HeaderView", "DetailView"):
                print(f"  Found '{key}' at {path}")
                if isinstance(val, dict):
                    for name, info in val.items():
                        if isinstance(info, dict):
                            lbl = info.get("Label", info.get("Text", info.get("Caption", "")))
                            style = info.get("Style", info.get("TabStyle", ""))
                            visible = info.get("Visible", "")
                            has_children = "Children" in info
                            print(f"    {name}: label='{lbl}' style='{style}' visible='{visible}' children={has_children}")
            if isinstance(val, dict):
                search_for_tabs(val, f"{path}.{key}", depth+1)

    search_for_tabs(form, "Form")

    # Specifically look for Acme-related controls
    print("\n=== SEARCHING FOR Acme/CBA/CHB CONTROLS ===")
    for term in ["Acme", "CBA", "CHB", "Consignment", "Asia"]:
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
                        ctrl_type = props.get("Type", "")
                        print(f"  [{term}] {name}: type='{ctrl_type}', label='{lbl}', value='{val}'")
        except:
            pass

    # Try opening known FastTab names that might exist
    print("\n=== TRYING TO OPEN FASTTAB SECTIONS ===")
    fasttab_candidates = [
        "AcmeAsia", "CBA_TabPage", "CBATabPage", "CBA_HeaderTab",
        "TabCBA", "CBA", "AcmeAsiaTab", "AcmeHuB", "CHB",
        "TabAcmeAsia", "CBACustomer", "CHBCustomer",
        "CreditAndCollections", "SalesOrderDefaults", "PaymentDefaults",
        "FinancialDimensions", "Warehouse", "InvoiceAndDelivery",
        "Transportation", "DirectDebitMandates", "Commerce",
        "TabGeneral", "General", "CreditCollections", "Credit",
        "HeaderView", "DetailView", "DetailsView",
        # Try tab page names that might contain CBA fields
        "TabHeaderGeneral", "TabHeaderCBA", "TabPageCBA",
        "CBAConsignment", "ConsignmentWarehouse",
    ]
    for tab_name in fasttab_candidates:
        try:
            tab_raw = client.call_tool("form_open_or_close_tab", {
                "tabName": tab_name, "tabAction": "Open"
            })
            if isinstance(tab_raw, str) and tab_raw.strip():
                tp = json.loads(tab_raw)
                fs = tp.get("FormState", {}).get("Form", {})
                result = tp.get("Result", "")
                if fs:
                    print(f"  {tab_name}: OPENED! Form keys: {sorted(fs.keys())[:10]}")
                    # Look for fields inside
                    detail = {"fields": {}, "grids": {}, "tabs": {}, "buttons": {}}
                    # Quick parse for fields
                    for ftype in ["Input", "Combobox", "Checkbox", "RealInput"]:
                        for fn in fs.get(ftype, {}):
                            fi = fs[ftype][fn]
                            if isinstance(fi, dict):
                                lbl = fi.get("Label", fn)
                                val = fi.get("Value", fi.get("ValueText", ""))
                                if val:
                                    print(f"    {fn}: label='{lbl}', value='{str(val)[:60]}'")
        except:
            pass

else:
    print("Empty response from row select!")

close_form(client)
print("\nDone.")
