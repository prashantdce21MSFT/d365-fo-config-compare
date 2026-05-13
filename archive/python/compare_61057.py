"""
Compare ADO 61057: Order Purpose: Sales for Free Beer (SG60)
Nav path: Accounts receivable > Setup > Acme Asia > Order purposes
Form: grid-based list of order purpose records

Reads from ENV1 (Env1 UAT Asia) and ENV4 (Env4 Config), compares by primary key.
"""
import sys
import json
import time
from datetime import datetime

sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365Region3Validator")

from d365_mcp_client import D365McpClient
from form_reader import (mcp_call, find_menu_item, open_form, close_form,
                         extract_form_data, _collect_all_grid_rows)

CONFIG_FILE = r"C:\D365DataValidator\config.json"
COMPANY = "sg60"
OUTPUT_DIR = r"C:\D365 Configuration Drift Analysis\output"


def load_config():
    with open(CONFIG_FILE) as f:
        return json.load(f)


def find_and_open_form(client, company, env_label):
    """Find the Order purposes menu item and open the form."""
    # First search for the menu item
    print(f"\n  [{env_label}] Searching for 'Order purpose' menu item...")
    matches = find_menu_item(client, "Order purpose", company_id=company)
    print(f"  [{env_label}] Found {len(matches)} matches:")
    for m in matches:
        print(f"    - {m['name']} ({m['type']}): {m['text']}")

    # Pick the right one — look for CBA order purposes
    menu_item = None
    for m in matches:
        if "orderpurpose" in m["name"].lower() or "cbaorderpurpose" in m["name"].lower():
            menu_item = m["name"]
            break
    if not menu_item and matches:
        menu_item = matches[0]["name"]

    if not menu_item:
        # Try broader search
        print(f"  [{env_label}] Trying broader search 'CBAOrder'...")
        matches = find_menu_item(client, "CBAOrder", company_id=company)
        for m in matches:
            print(f"    - {m['name']} ({m['type']}): {m['text']}")
            if "purpose" in m["name"].lower() or "purpose" in m["text"].lower():
                menu_item = m["name"]
                break
        if not menu_item and matches:
            menu_item = matches[0]["name"]

    if not menu_item:
        raise RuntimeError(f"Could not find Order purposes menu item on {env_label}")

    print(f"  [{env_label}] Opening form: {menu_item}")
    form_result = open_form(client, menu_item, menu_item_type="Display", company_id=company)

    form_state = form_result.get("FormState", {})
    print(f"  [{env_label}] Caption: {form_state.get('Caption', '?')}")
    print(f"  [{env_label}] Company: {form_state.get('Company', '?')}")

    return form_result, menu_item


def read_all_records(client, form_result, env_label):
    """Read all grid rows with pagination."""
    print(f"  [{env_label}] Reading grid records...")

    grid_name, all_rows, last_form = _collect_all_grid_rows(client, form_result)

    if grid_name:
        print(f"  [{env_label}] Grid '{grid_name}': {len(all_rows)} records")
    else:
        print(f"  [{env_label}] No grid found — checking for tab grids...")
        # Try extracting from tabs
        data = extract_form_data(client, form_result)
        for tab_name, tab_info in data.get("tabs", {}).items():
            for sg_name, sg_data in tab_info.get("grids", {}).items():
                rows = sg_data.get("rows", [])
                if rows:
                    print(f"  [{env_label}] Tab '{tab_name}' grid '{sg_name}': {len(rows)} rows")
                    all_rows.extend(rows)

    # Also extract top-level fields if present (parameter forms)
    data = extract_form_data(client, form_result)
    top_fields = data.get("fields", {})
    if top_fields:
        print(f"  [{env_label}] Top-level fields: {len(top_fields)}")

    return all_rows, top_fields


def identify_primary_key(rows):
    """Identify the primary key column(s) from grid rows."""
    if not rows:
        return []

    # Get all column names
    all_cols = set()
    for row in rows:
        all_cols.update(row.keys())

    # Heuristic: columns that have unique values across all rows
    candidates = []
    for col in sorted(all_cols):
        values = [row.get(col, "") for row in rows]
        non_empty = [v for v in values if v]
        if non_empty and len(set(non_empty)) == len(non_empty):
            candidates.append(col)

    # Prefer columns with "code", "id", "name", "purpose" in name
    priority_keywords = ["code", "id", "purpose", "name", "key", "number"]
    for kw in priority_keywords:
        for c in candidates:
            if kw in c.lower():
                return [c]

    # Fall back to first unique column
    if candidates:
        return [candidates[0]]

    # If no unique column, use all columns as composite key
    return sorted(all_cols)


def compare_records(rows_source, rows_target, pk_cols):
    """Compare records between source and target using primary key columns."""
    def make_key(row):
        return tuple(str(row.get(c, "")).strip() for c in pk_cols)

    source_by_key = {}
    for row in rows_source:
        key = make_key(row)
        source_by_key[key] = row

    target_by_key = {}
    for row in rows_target:
        key = make_key(row)
        target_by_key[key] = row

    all_keys = sorted(set(list(source_by_key.keys()) + list(target_by_key.keys())))

    results = []
    for key in all_keys:
        src = source_by_key.get(key)
        tgt = target_by_key.get(key)

        if src and tgt:
            # Both exist — compare field by field
            all_fields = sorted(set(list(src.keys()) + list(tgt.keys())))
            diffs = {}
            for f in all_fields:
                sv = str(src.get(f, "")).strip()
                tv = str(tgt.get(f, "")).strip()
                if sv != tv:
                    diffs[f] = {"source": sv, "target": tv}

            results.append({
                "key": dict(zip(pk_cols, key)),
                "status": "DIFF" if diffs else "MATCH",
                "diffs": diffs,
                "source": src,
                "target": tgt,
            })
        elif src and not tgt:
            results.append({
                "key": dict(zip(pk_cols, key)),
                "status": "SOURCE_ONLY",
                "diffs": {},
                "source": src,
                "target": None,
            })
        else:
            results.append({
                "key": dict(zip(pk_cols, key)),
                "status": "TARGET_ONLY",
                "diffs": {},
                "source": None,
                "target": tgt,
            })

    return results


def print_comparison(results, pk_cols):
    """Print comparison results in a readable table."""
    match_count = sum(1 for r in results if r["status"] == "MATCH")
    diff_count = sum(1 for r in results if r["status"] == "DIFF")
    source_only = sum(1 for r in results if r["status"] == "SOURCE_ONLY")
    target_only = sum(1 for r in results if r["status"] == "TARGET_ONLY")

    print(f"\n{'='*80}")
    print(f"  COMPARISON RESULTS")
    print(f"{'='*80}")
    print(f"  Primary Key: {', '.join(pk_cols)}")
    print(f"  Total records: {len(results)}")
    print(f"  MATCH:       {match_count}")
    print(f"  DIFF:        {diff_count}")
    print(f"  SOURCE_ONLY: {source_only} (in UAT but not in Config)")
    print(f"  TARGET_ONLY: {target_only} (in Config but not in UAT)")

    if diff_count > 0:
        print(f"\n--- DIFFERENCES ---")
        for r in results:
            if r["status"] == "DIFF":
                key_str = " | ".join(f"{k}={v}" for k, v in r["key"].items())
                print(f"\n  Key: {key_str}")
                for field, vals in r["diffs"].items():
                    print(f"    {field}:")
                    print(f"      UAT:    '{vals['source']}'")
                    print(f"      Config: '{vals['target']}'")

    if source_only > 0:
        print(f"\n--- SOURCE ONLY (UAT only) ---")
        for r in results:
            if r["status"] == "SOURCE_ONLY":
                key_str = " | ".join(f"{k}={v}" for k, v in r["key"].items())
                print(f"  {key_str}")
                if r["source"]:
                    for k, v in r["source"].items():
                        if v and k not in pk_cols:
                            print(f"    {k}: {v}")

    if target_only > 0:
        print(f"\n--- TARGET ONLY (Config only) ---")
        for r in results:
            if r["status"] == "TARGET_ONLY":
                key_str = " | ".join(f"{k}={v}" for k, v in r["key"].items())
                print(f"  {key_str}")
                if r["target"]:
                    for k, v in r["target"].items():
                        if v and k not in pk_cols:
                            print(f"    {k}: {v}")

    # Full detail table
    print(f"\n{'='*80}")
    print(f"  FULL RECORD COMPARISON TABLE")
    print(f"{'='*80}")

    if not results:
        print("  No records found.")
        return

    # Gather all field names
    all_fields = set()
    for r in results:
        if r["source"]:
            all_fields.update(r["source"].keys())
        if r["target"]:
            all_fields.update(r["target"].keys())
    all_fields = sorted(all_fields)

    # Print header
    print(f"\n  {'Status':<14} | ", end="")
    for pk in pk_cols:
        print(f"{pk:<20} | ", end="")
    other_fields = [f for f in all_fields if f not in pk_cols]
    for f in other_fields:
        print(f"{f:<30} | ", end="")
    print()
    print("  " + "-" * (14 + 3 + (23 * len(pk_cols)) + (33 * len(other_fields))))

    for r in results:
        src = r["source"] or {}
        tgt = r["target"] or {}
        status = r["status"]

        print(f"  {status:<14} | ", end="")
        for pk in pk_cols:
            val = r["key"].get(pk, "")
            print(f"{str(val):<20} | ", end="")

        for f in other_fields:
            sv = str(src.get(f, "")).strip()
            tv = str(tgt.get(f, "")).strip()
            if status == "SOURCE_ONLY":
                print(f"{sv:<30} | ", end="")
            elif status == "TARGET_ONLY":
                print(f"{tv:<30} | ", end="")
            elif sv == tv:
                print(f"{sv:<30} | ", end="")
            else:
                diff_str = f"UAT:{sv} | CFG:{tv}"
                print(f"{diff_str:<30} | ", end="")
        print()

    return results


def main():
    config = load_config()
    env1 = config["environments"]["ENV1"]  # Env1 UAT Asia (Source)
    env4 = config["environments"]["ENV4"]  # Env4 Config (Target)

    print(f"ADO 61057: Order Purpose - Sales for Free Beer (SG60)")
    print(f"Source: {env1['name']} ({env1['resource_url']})")
    print(f"Target: {env4['name']} ({env4['resource_url']})")
    print(f"Company: {COMPANY}")

    # Connect to both environments
    print(f"\nConnecting to environments...")
    client1 = D365McpClient(env1)
    client1.connect()
    print(f"  ENV1 ({env1['name']}): Connected")

    client4 = D365McpClient(env4)
    client4.connect()
    print(f"  ENV4 ({env4['name']}): Connected")

    # Read from Source (ENV1 - UAT)
    try:
        form_result1, mi1 = find_and_open_form(client1, COMPANY, env1["name"])
        rows1, fields1 = read_all_records(client1, form_result1, env1["name"])
        close_form(client1)
    except Exception as e:
        print(f"  ERROR reading from {env1['name']}: {e}")
        rows1, fields1 = [], {}

    time.sleep(1)

    # Read from Target (ENV4 - Config)
    try:
        form_result4, mi4 = find_and_open_form(client4, COMPANY, env4["name"])
        rows4, fields4 = read_all_records(client4, form_result4, env4["name"])
        close_form(client4)
    except Exception as e:
        print(f"  ERROR reading from {env4['name']}: {e}")
        rows4, fields4 = [], {}

    # Compare grid records
    if rows1 or rows4:
        print(f"\n  Source rows: {len(rows1)}")
        print(f"  Target rows: {len(rows4)}")

        # Identify primary key
        pk_cols = identify_primary_key(rows1 or rows4)
        print(f"  Identified PK: {pk_cols}")

        # Compare
        results = compare_records(rows1, rows4, pk_cols)
        print_comparison(results, pk_cols)

        # Save results
        import os
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        out_file = os.path.join(OUTPUT_DIR, "compare_61057_order_purposes.json")
        with open(out_file, "w", encoding="utf-8") as f:
            json.dump({
                "ado_id": 61057,
                "title": "Order Purpose: Sales for Free Beer (SG60)",
                "nav_path": "Accounts receivable > Setup > Acme Asia > Order purposes",
                "company": COMPANY,
                "source_env": env1["name"],
                "target_env": env4["name"],
                "source_rows": len(rows1),
                "target_rows": len(rows4),
                "pk_columns": pk_cols,
                "results": results,
                "timestamp": datetime.now().isoformat(),
            }, f, indent=2, ensure_ascii=False)
        print(f"\n  Results saved to: {out_file}")

    # Compare top-level fields if any
    if fields1 or fields4:
        print(f"\n--- Top-Level Fields Comparison ---")
        all_field_names = sorted(set(list(fields1.keys()) + list(fields4.keys())))
        for fn in all_field_names:
            f1 = fields1.get(fn, {})
            f4 = fields4.get(fn, {})
            v1 = str(f1.get("value", "")).strip()
            v4 = str(f4.get("value", "")).strip()
            label = f1.get("label", f4.get("label", fn))
            status = "MATCH" if v1 == v4 else "DIFF"
            if status == "DIFF":
                print(f"  DIFF  {label}: UAT='{v1}' | Config='{v4}'")
            else:
                print(f"  MATCH {label}: '{v1}'")


if __name__ == "__main__":
    main()
