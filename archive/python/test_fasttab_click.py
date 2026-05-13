"""Test: click on Acme section to expand FastTab, then read values."""
import sys, io, json, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient
from form_reader import open_form, close_form

with open(r"C:\D365DataValidator\config.json") as f:
    cfg = json.load(f)
client = D365McpClient(cfg["environments"]["ENV1"])
client.connect()

print("Opening CustTable on MY30...")
open_form(client, "CustTable", "Display", "MY30")

# First, search for all CBA/CHB-related controls to find section headers
print("\n=== ALL controls matching 'Acme' ===")
fc_raw = client.call_tool("form_find_controls", {"controlSearchTerm": "Acme"})
if isinstance(fc_raw, str) and fc_raw.strip():
    controls = json.loads(fc_raw)
    for c in controls:
        n = c.get("Name", "")
        p = c.get("Properties", {})
        print(f"  {n}: type={p.get('Type','')}, label='{p.get('Label','')}', value='{p.get('Value','')}'")
        # Print ALL property keys
        print(f"    Properties: {sorted(p.keys())}")

# Try to find the FastTab group/section control
print("\n=== Controls matching 'Tab' that might be FastTab headers ===")
for term in ["TabCBA", "TabCHB", "HeaderCBA", "HeaderCHB", "CBA_", "CHB_", "AcmeAsia", "AcmeHuB"]:
    fc_raw = client.call_tool("form_find_controls", {"controlSearchTerm": term})
    if isinstance(fc_raw, str) and fc_raw.strip():
        controls = json.loads(fc_raw)
        for c in controls:
            n = c.get("Name", "")
            p = c.get("Properties", {})
            print(f"  [{term}] {n}: label='{p.get('Label','')}', keys={sorted(p.keys())}")

# Select row 1
print("\n=== Selecting row 1 ===")
sel_raw = client.call_tool("form_select_grid_row", {
    "gridName": "Grid", "rowNumber": "1", "marking": "Unmarked"
})
if isinstance(sel_raw, str) and sel_raw.strip():
    sel = json.loads(sel_raw)
    form = sel.get("FormState", {}).get("Form", {})
    # Look for Acme tab in the response
    print(f"  Form keys: {sorted(form.keys())}")

    # Try clicking on "Acme" control to expand it
    print("\n=== Trying to click/expand Acme section ===")
    for ctrl_name in ["Acme", "AcmeAsia", "AcmeHuB", "CHB", "CBA"]:
        try:
            click_raw = client.call_tool("form_click_control", {
                "controlName": ctrl_name,
            })
            if isinstance(click_raw, str) and click_raw.strip():
                cr = json.loads(click_raw)
                cf = cr.get("FormState", {}).get("Form", {})
                if cf:
                    print(f"  Clicked '{ctrl_name}': got FormState! keys={sorted(cf.keys())}")
                    # Check for new tabs/fields
                    for key in ["Tab", "TabPage", "Input", "Combobox"]:
                        items = cf.get(key, {})
                        if items:
                            print(f"    {key}: {list(items.keys())[:10]}")
                else:
                    result = cr.get("Result", "")
                    if result:
                        print(f"  Clicked '{ctrl_name}': Result={str(result)[:100]}")
        except Exception as e:
            if "not found" not in str(e).lower():
                print(f"  Click '{ctrl_name}' error: {e}")

    # Try opening tab "Acme"
    print("\n=== Trying form_open_or_close_tab on Acme-related names ===")
    for tab_name in ["Acme", "AcmeAsia", "AcmeHuB", "CHBTab", "CBATab",
                      "CBA_TabPage", "CHB_TabPage", "TabAcme", "AcmeAsiaTab",
                      "CustTable_CHBConsignmentWarehouse"]:
        try:
            tab_raw = client.call_tool("form_open_or_close_tab", {
                "tabName": tab_name, "tabAction": "Open"
            })
            if isinstance(tab_raw, str) and tab_raw.strip():
                tp = json.loads(tab_raw)
                fs = tp.get("FormState", {}).get("Form", {})
                if fs:
                    print(f"  Opened '{tab_name}'! Form keys: {sorted(fs.keys())[:10]}")
                    # Look for our field
                    fc_raw2 = client.call_tool("form_find_controls", {"controlSearchTerm": "Consignment"})
                    if isinstance(fc_raw2, str) and fc_raw2.strip():
                        fc2 = json.loads(fc_raw2)
                        for c in fc2:
                            if "Consignment" in c.get("Name", ""):
                                v = c.get("Properties", {}).get("Value", "")
                                print(f"    After open: Consignment warehouse = '{v}'")
                                break
        except:
            pass

# Now select a DIFFERENT row and check
print("\n=== Selecting row 3, then checking Consignment ===")
sel_raw2 = client.call_tool("form_select_grid_row", {
    "gridName": "Grid", "rowNumber": "3", "marking": "Unmarked"
})

# Try: select row, then click "Acme" section, then read
for ctrl_name in ["Acme"]:
    try:
        click_raw = client.call_tool("form_click_control", {"controlName": ctrl_name})
    except:
        pass

fc_raw3 = client.call_tool("form_find_controls", {"controlSearchTerm": "Consignment"})
if isinstance(fc_raw3, str) and fc_raw3.strip():
    fc3 = json.loads(fc_raw3)
    for c in fc3:
        if "Consignment" in c.get("Name", ""):
            v = c.get("Properties", {}).get("Value", "")
            print(f"  Row 3 Consignment: '{v}'")

# Try: select row 5, open Acme, read
print("\n=== Select row 5, click Acme, read ===")
client.call_tool("form_select_grid_row", {
    "gridName": "Grid", "rowNumber": "5", "marking": "Unmarked"
})
try:
    client.call_tool("form_click_control", {"controlName": "Acme"})
except:
    pass

fc_raw4 = client.call_tool("form_find_controls", {"controlSearchTerm": "Consignment"})
if isinstance(fc_raw4, str) and fc_raw4.strip():
    fc4 = json.loads(fc_raw4)
    for c in fc4:
        if "Consignment" in c.get("Name", ""):
            v = c.get("Properties", {}).get("Value", "")
            print(f"  Row 5 Consignment: '{v}'")

close_form(client)
print("\nDone.")
