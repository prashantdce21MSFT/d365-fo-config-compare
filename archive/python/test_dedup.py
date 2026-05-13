"""Quick test: extract SG60 only from UAT to verify tab-specific field deduplication."""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from form_control_extractor import FormControlExtractor
from compare_form_controls import to_excel, collect_grid_data

legal_entities = ["SG60"]

print(">>> Extracting SG60 from UAT (ENV1)...")
ext = FormControlExtractor(env_key="ENV1")
result = ext.extract("InventLocations", legal_entities, "Display")

# Check detail field count and tab prefixes
for le in legal_entities:
    le_data = result["legal_entities"].get(le, {})
    for gname, gdata in le_data.get("grids", {}).items():
        print(f"\n{le} / {gname}:")
        print(f"  Rows: {gdata.get('total_rows', 0)}")
        rd = gdata.get("row_details", {})
        print(f"  Rows with details: {len(rd)}")
        if rd:
            # Check first row's fields
            first_key = list(rd.keys())[0]
            fields = rd[first_key]
            print(f"  Fields in row {first_key}: {len(fields)}")
            # Group by tab
            tabs = {}
            for fname, fdata in fields.items():
                tab = fdata.get("_tab", "(none)")
                if tab not in tabs:
                    tabs[tab] = []
                tabs[tab].append(fdata.get("label", fname))
            for tab_name in sorted(tabs.keys()):
                print(f"    [{tab_name}]: {len(tabs[tab_name])} fields")
                for f in sorted(tabs[tab_name]):
                    print(f"      - {f}")

# Build grid data to check for duplicates
env_results = [("UAT", result)]
grids_info = collect_grid_data(env_results, legal_entities)
for gname, ginfo in grids_info.items():
    print(f"\nGrid '{gname}' detail labels ({len(ginfo['detail_labels'])} total):")
    for dl in sorted(ginfo["detail_labels"]):
        print(f"  {dl}")

# Write Excel
env_results_full = [("UAT", result), ("Config", result)]  # use same data for both to test layout
filepath = to_excel(env_results_full, legal_entities, result, result, ado_id="TEST", area_path="TEST")
print(f"\nDone: {filepath}")
