"""
Read CBAInvoiceChargeTrackingParameter form from ENV1 (Env1 UAT Asia) and ENV4 (Env4 Config).
Company: sg60
"""
import sys
import json
import time

sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient
from form_reader import mcp_call, open_form, close_form, extract_form_data, _extract_fields

CONFIG_FILE = r"C:\D365DataValidator\config.json"
MENU_ITEM = "CBAInvoiceChargeTrackingParameter"
COMPANY = "sg60"


def load_config():
    with open(CONFIG_FILE) as f:
        return json.load(f)


def read_form_all_tabs(client, env_label):
    """Open the form, read all fields including all tabs."""
    print(f"\n{'='*70}")
    print(f"  READING: {MENU_ITEM} | Company: {COMPANY} | Env: {env_label}")
    print(f"{'='*70}")

    # Open the form
    print(f"  Opening form...")
    form_result = open_form(client, MENU_ITEM, menu_item_type="Display", company_id=COMPANY)

    form_state = form_result.get("FormState", {})
    form = form_state.get("Form", {})
    print(f"  Caption: {form_state.get('Caption', '?')}")
    print(f"  Company: {form_state.get('Company', '?')}")

    # Extract initial data (gets the active tab's fields)
    data = extract_form_data(client, form_result)
    print(f"  Tabs found: {list(data['tabs'].keys())}")

    # Check for closed tabs and try to activate them
    for tab_name, tab_info in list(data["tabs"].items()):
        tab_text = tab_info.get("text", tab_name)
        has_content = bool(tab_info.get("fields")) or bool(tab_info.get("grids"))
        if has_content:
            print(f"  Tab '{tab_text}' ({tab_name}): {len(tab_info.get('fields', {}))} fields (loaded)")
            continue

        # Tab is closed/empty - try to activate it
        print(f"  Tab '{tab_text}' ({tab_name}): empty, trying to activate...")

        # Method: controlName="Tab", actionId=<tab_page_name>
        try:
            result = mcp_call(client, "form_click_control", {
                "controlName": "Tab",
                "actionId": tab_name,
            })
            if "raw" not in result:
                new_form = result.get("FormState", {}).get("Form", {})
                # Check for the tab's content
                new_tab = new_form.get("Tab", {}).get(tab_name, {})
                children = new_tab.get("Children", {})
                if isinstance(children, dict) and children:
                    # Re-extract data from the updated form
                    new_data = extract_form_data(client, result)
                    for tn, ti in new_data["tabs"].items():
                        if ti.get("fields") or ti.get("grids"):
                            data["tabs"][tn] = ti
                    print(f"    Activated! Got {len(new_data['tabs'].get(tab_name, {}).get('fields', {}))} fields")
                else:
                    # Tab activated but still no Children - may have grid/fields at top level
                    new_data = extract_form_data(client, result)
                    # Check if any new fields/grids appeared
                    for tn, ti in new_data["tabs"].items():
                        if ti.get("fields") or ti.get("grids"):
                            data["tabs"][tn] = ti
                    # Also check top-level grids
                    if new_data["grids"]:
                        # Assign top-level grids to this tab
                        data["tabs"][tab_name]["grids"] = new_data["grids"]
                        print(f"    Got top-level grids: {list(new_data['grids'].keys())}")
                    # Check for new top-level fields
                    for fn, fi in new_data["fields"].items():
                        if fn not in data["fields"]:
                            data["fields"][fn] = fi
                    new_field_count = len(new_data["fields"]) + sum(
                        len(ti.get("fields", {})) for ti in new_data["tabs"].values()
                    )
                    print(f"    Tab switch response: {new_field_count} total fields across all tabs")

                    # Dump the full raw response to understand its structure
                    ns_tab_raw = new_form.get("Tab", {}).get(tab_name, {})
                    print(f"    Raw tab after switch: {json.dumps(ns_tab_raw, indent=2, default=str)[:500]}")
            else:
                print(f"    Failed: {result.get('raw', '')[:200]}")
        except Exception as e:
            print(f"    Error: {e}")

    close_form(client)
    return data


def print_env_data(data, env_label):
    """Print all field values from the form data."""
    print(f"\n{'='*70}")
    print(f"  RESULTS FOR: {env_label}")
    print(f"{'='*70}")

    if data["fields"]:
        print(f"\n  --- Top-level Fields ---")
        for fname, finfo in sorted(data["fields"].items()):
            label = finfo.get("label", fname)
            value = finfo.get("value", "")
            ftype = finfo.get("type", "?")
            print(f"    [{ftype:>8}] {label}: {value}")

    for tab_name, tab_info in sorted(data["tabs"].items()):
        tab_text = tab_info.get("text", tab_name)
        fields = tab_info.get("fields", {})
        grids = tab_info.get("grids", {})

        if fields or grids:
            print(f"\n  --- Tab: {tab_text} ---")
            for fname, finfo in sorted(fields.items()):
                label = finfo.get("label", fname)
                value = finfo.get("value", "")
                ftype = finfo.get("type", "?")
                print(f"    [{ftype:>8}] {label}: {value}")

            for grid_name, grid_info in grids.items():
                rows = grid_info.get("rows", [])
                cols = grid_info.get("columns", [])
                print(f"\n    Grid: {grid_name} ({len(rows)} rows)")
                if cols:
                    print(f"    Columns: {cols}")
                for i, row in enumerate(rows):
                    print(f"      Row {i+1}: {row}")
        else:
            print(f"\n  --- Tab: {tab_text} --- (no data available)")


def main():
    config = load_config()

    env1_config = config["environments"]["ENV1"]
    env4_config = config["environments"]["ENV4"]

    # ENV1: Env1 UAT Asia
    print("\n" + "=" * 70)
    print("  CONNECTING TO ENV1: Env1 UAT Asia")
    print("=" * 70)
    client1 = D365McpClient(env1_config)
    client1.connect()
    data1 = read_form_all_tabs(client1, "ENV1 - Env1 UAT Asia")

    # ENV4: Env4 Config
    print("\n" + "=" * 70)
    print("  CONNECTING TO ENV4: Env4 Config")
    print("=" * 70)
    client4 = D365McpClient(env4_config)
    client4.connect()
    data4 = read_form_all_tabs(client4, "ENV4 - Env4 Config")

    # Print results
    print_env_data(data1, "ENV1 - Env1 UAT Asia")
    print_env_data(data4, "ENV4 - Env4 Config")

    # Side-by-side comparison
    print(f"\n{'='*70}")
    print(f"  SIDE-BY-SIDE COMPARISON")
    print(f"{'='*70}")

    def flatten_fields(data):
        flat = {}
        for fname, finfo in data["fields"].items():
            label = finfo.get("label", fname)
            flat[label] = finfo.get("value", "")
        for tab_name, tab_info in data["tabs"].items():
            tab_text = tab_info.get("text", tab_name)
            for fname, finfo in tab_info.get("fields", {}).items():
                label = finfo.get("label", fname)
                key = f"[{tab_text}] {label}"
                flat[key] = finfo.get("value", "")
        return flat

    flat1 = flatten_fields(data1)
    flat4 = flatten_fields(data4)
    all_keys = sorted(set(flat1.keys()) | set(flat4.keys()))

    print(f"\n  {'Field':<60} {'ENV1 (UAT)':<30} {'ENV4 (Config)':<30} {'Match?'}")
    print(f"  {'-'*60} {'-'*30} {'-'*30} {'-'*6}")
    for key in all_keys:
        v1 = flat1.get(key, "<missing>")
        v4 = flat4.get(key, "<missing>")
        match = "YES" if v1 == v4 else "DIFF"
        print(f"  {key:<60} {str(v1):<30} {str(v4):<30} {match}")


if __name__ == "__main__":
    main()
