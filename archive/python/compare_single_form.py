"""Compare a single form across UAT (ENV1) and Config (ENV4)."""
import sys, os, json
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient
from form_reader import (find_menu_item, open_form, close_form,
                         read_form_records, resolve_menu_item_from_path)

CONFIG_FILE = r"C:\D365DataValidator\config.json"

with open(CONFIG_FILE) as f:
    config = json.load(f)

TITLE = "Sales and marketing > Setup > Sales and marketing parameters"
COMPANY = "MY60"

# Connect to both environments
print("Connecting to UAT (ENV1)...")
src = D365McpClient(config["environments"]["ENV1"])
src.connect()
print("  OK")

print("Connecting to Config (ENV4)...")
tgt = D365McpClient(config["environments"]["ENV4"])
tgt.connect()
print("  OK")

# Resolve menu item from path
print(f"\nResolving menu item for: {TITLE}")
mi_name, mi_type, hint = resolve_menu_item_from_path(src, TITLE, COMPANY)
print(f"  Found: mi={mi_name} type={mi_type} hint={hint}")

if not mi_name:
    # Manual search
    print("  Trying manual search for 'Order purposes'...")
    items = find_menu_item(src, "Order purposes", COMPANY)
    for it in items:
        print(f"    {it['name']} ({it['type']}): {it['text']}")
    if items:
        mi_name = items[0]["name"]
        mi_type = items[0]["type"]

if not mi_name:
    print("ERROR: Could not find menu item")
    sys.exit(1)

# Read from UAT
print(f"\n--- UAT (ENV1) ---")
print(f"Opening {mi_name} on {COMPANY}...")
close_form(src)
form_result = open_form(src, mi_name, mi_type, COMPANY)
src_records, src_cols = read_form_records(src, form_result, include_details=False)
print(f"  Records: {len(src_records)}")
close_form(src)

# Read from Config
print(f"\n--- Config (ENV4) ---")
print(f"Opening {mi_name} on {COMPANY}...")
close_form(tgt)
form_result = open_form(tgt, mi_name, mi_type, COMPANY)
tgt_records, tgt_cols = read_form_records(tgt, form_result, include_details=False)
print(f"  Records: {len(tgt_records)}")
close_form(tgt)

# Compare
print(f"\n{'='*70}")
print(f"COMPARISON: {TITLE}")
print(f"Company: {COMPANY}  |  Menu Item: {mi_name}")
print(f"UAT records: {len(src_records)}  |  Config records: {len(tgt_records)}")
print(f"{'='*70}")

# Find key column (first column)
all_cols = sorted(set(src_cols + tgt_cols))
print(f"Columns: {all_cols}")

# Build lookup by first meaningful column
def make_key(record, cols):
    """Use first 2 non-empty columns as key."""
    vals = []
    for c in sorted(record.keys()):
        v = str(record.get(c, "")).strip()
        if v:
            vals.append(v)
        if len(vals) >= 2:
            break
    return "|".join(vals)

src_map = {}
for r in src_records:
    k = make_key(r, all_cols)
    src_map[k] = r

tgt_map = {}
for r in tgt_records:
    k = make_key(r, all_cols)
    tgt_map[k] = r

all_keys = sorted(set(list(src_map.keys()) + list(tgt_map.keys())))

matches = 0
diffs = 0
only_src = 0
only_tgt = 0

print(f"\n--- MATCHES ---")
for k in all_keys:
    s = src_map.get(k)
    t = tgt_map.get(k)
    if s and t:
        # Compare field by field
        diff_fields = []
        for c in all_cols:
            sv = str(s.get(c, "")).strip()
            tv = str(t.get(c, "")).strip()
            if sv != tv:
                diff_fields.append((c, sv, tv))
        if diff_fields:
            diffs += 1
            print(f"\n  DIFF [{k}]:")
            for c, sv, tv in diff_fields:
                print(f"    {c}: UAT='{sv}' vs Config='{tv}'")
        else:
            matches += 1
    elif s and not t:
        only_src += 1
        print(f"\n  ONLY IN UAT: [{k}]")
        for c in all_cols:
            v = str(s.get(c, "")).strip()
            if v:
                print(f"    {c}: {v}")
    else:
        only_tgt += 1
        print(f"\n  ONLY IN CONFIG: [{k}]")
        for c in all_cols:
            v = str(t.get(c, "")).strip()
            if v:
                print(f"    {c}: {v}")

print(f"\n{'='*70}")
print(f"SUMMARY: {matches} match, {diffs} diff, {only_src} only-UAT, {only_tgt} only-Config")
print(f"{'='*70}")
