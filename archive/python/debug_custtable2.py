"""Debug: use FormControlExtractor to open CustTable, then inspect tabs."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from form_control_extractor import FormControlExtractor

ext = FormControlExtractor(env_key="ENV1")

# Use extract but only for MY30, just to get the client and form open
# We'll intercept after extraction to check what tabs were found
result = ext.extract("CustTable", ["MY30"], "Display")

# Check what was captured
le_data = result["legal_entities"].get("MY30", {})
for gname, gdata in le_data.get("grids", {}).items():
    print(f"\nGrid '{gname}': {gdata.get('total_rows', 0)} rows")
    rd = gdata.get("row_details", {})
    print(f"  Rows with details: {len(rd)}")
    if rd:
        first_key = list(rd.keys())[0]
        fields = rd[first_key]
        print(f"  Fields in first row: {len(fields)}")
        tabs = {}
        for fname, fdata in fields.items():
            tab = fdata.get("_tab", "(none)")
            if tab not in tabs:
                tabs[tab] = []
            tabs[tab].append((fdata.get("label", fname), fdata.get("value", "")))
        for tab_name in sorted(tabs.keys()):
            print(f"\n  [{tab_name}] ({len(tabs[tab_name])} fields):")
            for lbl, val in sorted(tabs[tab_name]):
                print(f"    {lbl} = {val[:60] if val else ''}")

# Check if "Consignment" appears anywhere
print("\n=== SEARCHING ALL ROWS FOR 'consignment' ===")
for ri, fields in rd.items():
    for fname, fdata in fields.items():
        lbl = str(fdata.get("label", "")).lower()
        name_lower = fname.lower()
        if "consignment" in lbl or "consignment" in name_lower:
            val = fdata.get("value", "")
            tab = fdata.get("_tab", "")
            print(f"  Row {ri}: {fname} label='{fdata.get('label','')}' value='{val}' tab='{tab}'")
            break

print("\nDone.")
