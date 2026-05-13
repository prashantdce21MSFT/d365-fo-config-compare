"""Find all CBA/CHB fields in CustomersV3 entity."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"C:\D365DataValidator")

from d365_mcp_client import D365McpClient

with open(r"C:\D365DataValidator\config.json") as f:
    cfg = json.load(f)
client = D365McpClient(cfg["environments"]["ENV1"])
client.connect()

meta = client.get_entity_metadata("CustomersV3")
parsed = json.loads(meta) if isinstance(meta, str) else meta
fields = parsed.get("Properties", {})

print("=== CBA/CHB/Acme fields in CustomersV3 ===")
for fname in sorted(fields.keys()):
    if any(t in fname for t in ["CBA", "CHB", "Acme"]):
        finfo = fields[fname]
        print(f"  {fname}: {finfo}")

print(f"\n=== Total fields: {len(fields)} ===")

# Also query one customer to see actual CBA/CHB values
print("\n=== Customer 210205 MY30 — CBA/CHB fields ===")
cba_fields = [f for f in fields if any(t in f for t in ["CBA", "CHB"])]
select = ",".join(["CustomerAccount", "dataAreaId"] + cba_fields[:30])
raw = client.find_entities("CustomersV3", f"$filter=CustomerAccount eq '210205' and dataAreaId eq 'MY30'&$select={select}&$top=1")
if raw:
    data = json.loads(raw) if isinstance(raw, str) else raw
    for rec in data.get("value", []):
        for k, v in sorted(rec.items()):
            if v and str(v).strip() and not k.startswith("@"):
                print(f"  {k} = {v}")

print("\nDone.")
