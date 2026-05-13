"""Quick test: CustTable extraction with FastTab fix, MY30 only, 5 rows."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from form_control_extractor import FormControlExtractor

ext = FormControlExtractor(env_key="ENV1")
result = ext.extract("CustTable", ["MY30"], "Display")

le_data = result["legal_entities"].get("MY30", {})
for gname, gdata in le_data.get("grids", {}).items():
    print(f"\nGrid '{gname}': {gdata.get('total_rows', 0)} rows")
    rd = gdata.get("row_details", {})
    print(f"  Rows with details: {len(rd)}")
    if rd:
        # Show first 3 rows
        for ri in sorted(rd.keys(), key=lambda x: int(x))[:3]:
            fields = rd[ri]
            print(f"\n  Row {ri}: {len(fields)} fields")
            tabs = {}
            for fname, fdata in fields.items():
                tab = fdata.get("_tab", "(none)")
                if tab not in tabs:
                    tabs[tab] = []
                tabs[tab].append((fdata.get("label", fname), fdata.get("value", "")))
            for tab_name in sorted(tabs.keys()):
                print(f"    [{tab_name}] ({len(tabs[tab_name])} fields):")
                for lbl, val in sorted(tabs[tab_name])[:5]:
                    print(f"      {lbl} = {val[:60] if val else ''}")
                if len(tabs[tab_name]) > 5:
                    print(f"      ... +{len(tabs[tab_name])-5} more")

        # Check specifically for consignment across all rows
        print("\n=== CONSIGNMENT FIELD ACROSS ROWS ===")
        for ri in sorted(rd.keys(), key=lambda x: int(x))[:10]:
            for fname, fdata in rd[ri].items():
                if "consignment" in fname.lower() or "consignment" in str(fdata.get("label", "")).lower():
                    print(f"  Row {ri}: {fdata.get('label','')} = '{fdata.get('value','')}'")
                    break

print("\nDone.")
