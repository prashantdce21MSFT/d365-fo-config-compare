"""Query CustTable via OData with correct parameters."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"C:\D365DataValidator")

from d365_mcp_client import D365McpClient

with open(r"C:\D365DataValidator\config.json") as f:
    cfg = json.load(f)

# UAT
client = D365McpClient(cfg["environments"]["ENV1"])
client.connect()

# Step 1: Find entity type
print("=== Finding entity for CustTable ===")
raw = client.find_entity_types("CustTable")
print(f"  {raw[:500] if raw else 'empty'}")

# Step 2: Try CustomersV3 entity
print("\n=== CustomersV3 metadata ===")
meta = client.get_entity_metadata("CustomersV3")
if meta:
    parsed = json.loads(meta) if isinstance(meta, str) else meta
    # Look for Consignment/warehouse fields
    if isinstance(parsed, dict):
        for key in ["Properties", "Fields", "properties", "fields"]:
            if key in parsed:
                fields = parsed[key]
                if isinstance(fields, list):
                    matches = [f for f in fields if any(t in str(f).lower() for t in ["consignment", "chb", "warehouse"])]
                    print(f"  Matching fields: {matches[:10]}")
                    print(f"  Total fields: {len(fields)}")
                elif isinstance(fields, dict):
                    matches = {k: v for k, v in fields.items() if any(t in k.lower() for t in ["consignment", "chb", "warehouse"])}
                    print(f"  Matching fields: {matches}")
                    print(f"  Total fields: {len(fields)}")
    else:
        print(f"  Type: {type(parsed)}, content: {str(parsed)[:400]}")

# Step 3: Query customer 210205 in MY30
print("\n=== Query customer 210205 MY30 ===")
try:
    raw3 = client.find_entities(
        "CustomersV3",
        "$filter=CustomerAccount eq '210205' and dataAreaId eq 'MY30'&$select=CustomerAccount,OrganizationName,CHBConsignmentWarehouse,WarehouseId&$top=1"
    )
    if raw3:
        print(f"  {raw3[:800]}")
except Exception as e:
    print(f"  Error: {e}")

# Try without CHB prefix
print("\n=== Query with different field names ===")
try:
    raw4 = client.find_entities(
        "CustomersV3",
        "$filter=CustomerAccount eq '210205' and dataAreaId eq 'MY30'&$top=1"
    )
    if raw4:
        data = json.loads(raw4) if isinstance(raw4, str) else raw4
        if isinstance(data, dict) and "value" in data:
            for rec in data["value"][:1]:
                # Print fields with warehouse/consignment
                for k, v in sorted(rec.items()):
                    if any(t in k.lower() for t in ["consignment", "warehouse", "chb", "account", "name"]):
                        print(f"  {k} = {v}")
        else:
            print(f"  {str(data)[:800]}")
except Exception as e:
    print(f"  Error: {e}")

# Query customer 1000 in MY60
print("\n=== Query customer 1000 MY60 ===")
try:
    raw5 = client.find_entities(
        "CustomersV3",
        "$filter=CustomerAccount eq '1000' and dataAreaId eq 'MY60'&$top=1"
    )
    if raw5:
        data5 = json.loads(raw5) if isinstance(raw5, str) else raw5
        if isinstance(data5, dict) and "value" in data5:
            for rec in data5["value"][:1]:
                for k, v in sorted(rec.items()):
                    if any(t in k.lower() for t in ["consignment", "warehouse", "chb", "account", "name"]):
                        print(f"  {k} = {v}")
        else:
            print(f"  {str(data5)[:800]}")
except Exception as e:
    print(f"  Error: {e}")

print("\nDone.")
