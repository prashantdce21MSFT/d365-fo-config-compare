"""Test: take LAST non-empty value from form_find_controls duplicates."""
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

# Open on MY60 where user showed 9099_1 for customer 1000
print("Opening CustTable on MY60...")
open_form(client, "CustTable", "Display", "MY60")

# Select several rows and check ALL copies of Consignment warehouse
for row_num in range(1, 11):
    sel_raw = client.call_tool("form_select_grid_row", {
        "gridName": "Grid", "rowNumber": str(row_num), "marking": "Unmarked"
    })

    # Get customer account from grid
    acct = "?"
    if isinstance(sel_raw, str) and sel_raw.strip():
        sel = json.loads(sel_raw)
        form = sel.get("FormState", {}).get("Form", {})
        # Navigate to grid rows
        tab = form.get("Tab", {}).get("TabPageGrid", {})
        children = tab.get("Children", tab)
        for gtype in ["Grid"]:
            for gn, gi in children.get(gtype, {}).items():
                if isinstance(gi, dict):
                    for r in gi.get("Rows", []):
                        v = r.get("Values", {})
                        a = v.get("Account", v.get("Customer account", ""))
                        if a:
                            acct = a
                            break

    # Get ALL consignment warehouse results
    fc_raw = client.call_tool("form_find_controls", {"controlSearchTerm": "CustTable_CHBConsignment"})
    vals = []
    if isinstance(fc_raw, str) and fc_raw.strip():
        controls = json.loads(fc_raw)
        for c in controls:
            if "Consignment" in c.get("Name", ""):
                v = c.get("Properties", {}).get("Value", "")
                vals.append(v)

    # Also try the more specific search
    fc_raw2 = client.call_tool("form_find_controls", {"controlSearchTerm": "Consignment"})
    vals2 = []
    if isinstance(fc_raw2, str) and fc_raw2.strip():
        controls2 = json.loads(fc_raw2)
        for c in controls2:
            if "Consignment" in c.get("Name", ""):
                v = c.get("Properties", {}).get("Value", "")
                vals2.append(v)

    print(f"Row {row_num:2d} ({acct:>8s}): specific={vals}, broad={vals2}")
    time.sleep(0.2)

close_form(client)
print("\nDone.")
