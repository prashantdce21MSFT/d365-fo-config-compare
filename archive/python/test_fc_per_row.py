"""Test: does form_find_controls update per row selection?"""
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

# Select rows 1-5 and check Consignment warehouse each time
for row_num in [1, 2, 3, 4, 5]:
    print(f"\n--- Selecting row {row_num} ---")
    sel_raw = client.call_tool("form_select_grid_row", {
        "gridName": "Grid", "rowNumber": str(row_num), "marking": "Unmarked"
    })
    if isinstance(sel_raw, str) and sel_raw.strip():
        sel = json.loads(sel_raw)
        form = sel.get("FormState", {}).get("Form", {})
        # Get grid row data to see which customer this is
        grid = form.get("Tab", {}).get("TabPageGrid", {}).get("Children", {}).get("Grid", {}).get("Grid", {})
        if not grid:
            # Try nested structure
            for key in ["Tab", "TabPage"]:
                for tname, tinfo in form.get(key, {}).items():
                    children = tinfo.get("Children", tinfo)
                    if isinstance(children, dict):
                        for gtype in ["Grid"]:
                            for gn, gi in children.get(gtype, {}).items():
                                if isinstance(gi, dict) and gi.get("Rows"):
                                    grid = gi

        rows = grid.get("Rows", []) if isinstance(grid, dict) else []
        if rows:
            vals = rows[0].get("Values", {}) if rows else {}
            acct = vals.get("Account", vals.get("Customer account", "?"))
            name = vals.get("Name", "?")
            print(f"  Customer: {acct} - {name}")

        # Now check form_find_controls for Consignment
        fc_raw = client.call_tool("form_find_controls", {"controlSearchTerm": "Consignment"})
        if isinstance(fc_raw, str) and fc_raw.strip():
            controls = json.loads(fc_raw)
            for c in controls:
                cname = c.get("Name", "")
                if "Consignment" in cname or "consignment" in cname.lower():
                    val = c.get("Properties", {}).get("Value", "")
                    lbl = c.get("Properties", {}).get("Label", "")
                    print(f"  {cname}: label='{lbl}', value='{val}'")

        # Also check Warehouse and Credit limit
        fc_raw2 = client.call_tool("form_find_controls", {"controlSearchTerm": "Warehouse"})
        if isinstance(fc_raw2, str) and fc_raw2.strip():
            controls2 = json.loads(fc_raw2)
            for c in controls2:
                cname = c.get("Name", "")
                props = c.get("Properties", {})
                if cname == "CustTable_InventLocation" or "Warehouse" in props.get("Label", ""):
                    val = props.get("Value", "")
                    lbl = props.get("Label", "")
                    print(f"  {cname}: label='{lbl}', value='{val}'")
    else:
        print("  Empty response")
    time.sleep(0.3)

close_form(client)
print("\nDone.")
