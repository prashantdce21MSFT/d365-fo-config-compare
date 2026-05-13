import sys, json
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")
from d365_mcp_client import D365McpClient
from form_reader import open_form, close_form, read_form_records

with open(r"C:\D365DataValidator\config.json") as f:
    config = json.load(f)

MI = "CBAOrderPurposes"
COMPANY = "SG60"

src = D365McpClient(config["environments"]["ENV1"])
src.connect()
print("UAT connected")

tgt = D365McpClient(config["environments"]["ENV4"])
tgt.connect()
print("Config connected")

# UAT
print(f"\n--- UAT ---")
close_form(src)
fr = open_form(src, MI, "Display", COMPANY)
src_recs, src_cols = read_form_records(src, fr, include_details=False)
print(f"  {len(src_recs)} records, cols: {src_cols}")
for r in src_recs:
    print(f"  {r}")
close_form(src)

# Config
print(f"\n--- Config ---")
close_form(tgt)
fr = open_form(tgt, MI, "Display", COMPANY)
tgt_recs, tgt_cols = read_form_records(tgt, fr, include_details=False)
print(f"  {len(tgt_recs)} records, cols: {tgt_cols}")
for r in tgt_recs:
    print(f"  {r}")
close_form(tgt)

# Compare using first column as key
print(f"\n{'='*70}")
print(f"COMPARISON: Order purposes (SG60)")
print(f"UAT: {len(src_recs)} | Config: {len(tgt_recs)}")
print(f"{'='*70}")

# Use first column as key
key_col = src_cols[0] if src_cols else None
if not key_col and tgt_cols:
    key_col = tgt_cols[0]

all_cols = sorted(set(src_cols + tgt_cols))

src_map = {str(r.get(key_col, "")).strip(): r for r in src_recs} if key_col else {}
tgt_map = {str(r.get(key_col, "")).strip(): r for r in tgt_recs} if key_col else {}

all_keys = sorted(set(list(src_map.keys()) + list(tgt_map.keys())))

matches = diffs = only_src = only_tgt = 0
for k in all_keys:
    s = src_map.get(k)
    t = tgt_map.get(k)
    if s and t:
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
    elif s:
        only_src += 1
        vals = " | ".join(f"{c}={s.get(c,'')}" for c in all_cols if s.get(c))
        print(f"  ONLY UAT: {vals}")
    else:
        only_tgt += 1
        vals = " | ".join(f"{c}={t.get(c,'')}" for c in all_cols if t.get(c))
        print(f"  ONLY CONFIG: {vals}")

print(f"\nSUMMARY: {matches} match, {diffs} diff, {only_src} only-UAT, {only_tgt} only-Config")
