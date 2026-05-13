"""Read CBAInvoiceChargeTrackingParameter for all companies on both envs."""
import sys, json, time
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")
from d365_mcp_client import D365McpClient
from form_reader import open_form, extract_form_data, close_form, mcp_call

CONFIG_FILE = r"C:\D365DataValidator\config.json"
MENU_ITEM = "CBAInvoiceChargeTrackingParameter"
COMPANIES = ["my30", "sg60", "my60", "dat"]

with open(CONFIG_FILE) as f:
    config = json.load(f)

src = D365McpClient(config["environments"]["ENV1"])
tgt = D365McpClient(config["environments"]["ENV4"])
src.connect()
print("[OK] Env1 UAT Asia (source)")
tgt.connect()
print("[OK] Env4 Config (target)")

def read_form_data(client, env_name, mi, company):
    """Open form and read all fields on all tabs."""
    print(f"\n  --- {env_name} | cmp={company} ---")
    try:
        close_form(client)
        time.sleep(0.5)
        form_result = open_form(client, mi, company_id=company)
        data = extract_form_data(client, form_result)
        
        if not data:
            print("    No data returned")
            return None
        
        print(f"    Caption: {data.get('caption', '?')}")
        print(f"    Company: {data.get('company', '?')}")
        
        # Print top-level fields
        if data.get("fields"):
            print(f"    Top-level fields ({len(data['fields'])}):")
            for k, v in sorted(data["fields"].items()):
                val = v.get("value", "") if isinstance(v, dict) else v
                label = v.get("label", k) if isinstance(v, dict) else k
                if val:
                    print(f"      {label}: {val}")
        
        # Print tab fields
        for tab_name, tab_data in data.get("tabs", {}).items():
            tab_text = tab_data.get("text", tab_name)
            fields = tab_data.get("fields", {})
            grids = tab_data.get("grids", {})
            if fields or grids:
                print(f"    Tab '{tab_text}':")
                for k, v in sorted(fields.items()):
                    val = v.get("value", "") if isinstance(v, dict) else v
                    label = v.get("label", k) if isinstance(v, dict) else k
                    if val:
                        print(f"      {label}: {val}")
                for gname, gdata in grids.items():
                    rows = gdata.get("rows", [])
                    print(f"      Grid '{gname}': {len(rows)} rows")
                    for row in rows[:5]:
                        print(f"        {row}")
        
        # Print grids
        for gname, gdata in data.get("grids", {}).items():
            rows = gdata.get("rows", [])
            print(f"    Grid '{gname}': {len(rows)} rows")
            for row in rows[:5]:
                print(f"      {row}")
        
        # Try switching to closed tabs via form_click_control
        for tab_name, tab_data in data.get("tabs", {}).items():
            tab_text = tab_data.get("text", tab_name)
            if not tab_data.get("fields") and not tab_data.get("grids"):
                print(f"    Trying to activate tab '{tab_text}'...")
                try:
                    click_result = mcp_call(client, "form_click_control", {"controlName": tab_name})
                    time.sleep(1)
                    # Re-read form state
                    state = mcp_call(client, "form_get_controls", {})
                    re_data = extract_form_data(client, state) if isinstance(state, dict) else None
                    if re_data:
                        for tn, td in re_data.get("tabs", {}).items():
                            if td.get("fields"):
                                print(f"    Tab '{tn}' (after click):")
                                for k, v in sorted(td["fields"].items()):
                                    val = v.get("value", "") if isinstance(v, dict) else v
                                    label = v.get("label", k) if isinstance(v, dict) else k
                                    if val:
                                        print(f"      {label}: {val}")
                except Exception as e:
                    print(f"    Tab activate failed: {e}")

        close_form(client)
        return data
        
    except Exception as e:
        print(f"    Error: {e}")
        try:
            close_form(client)
        except:
            pass
        return None

print(f"\nReading {MENU_ITEM} for companies: {', '.join(COMPANIES)}")
print("=" * 70)

for company in COMPANIES:
    print(f"\n{'='*70}")
    print(f"  COMPANY: {company.upper()}")
    print(f"{'='*70}")
    
    src_data = read_form_data(src, "Env1 UAT Asia", MENU_ITEM, company)
    time.sleep(1)
    tgt_data = read_form_data(tgt, "Env4 Config", MENU_ITEM, company)
    time.sleep(1)

print("\nDone.")
