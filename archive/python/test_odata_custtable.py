"""Query CustTable via OData to get Consignment warehouse per customer."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient

with open(r"C:\D365DataValidator\config.json") as f:
    cfg = json.load(f)
client = D365McpClient(cfg["environments"]["ENV1"])
client.connect()

# Step 1: Find the entity type for CustTable
print("=== Finding entity type for CustTable ===")
raw = client.call_tool("data_find_entity_type", {"searchTerm": "CustTable"})
if isinstance(raw, str) and raw.strip():
    result = json.loads(raw)
    if isinstance(result, list):
        for ent in result[:5]:
            print(f"  {ent.get('Name', ent.get('name', ''))}: {ent}")
    else:
        print(f"  Result: {str(result)[:500]}")

# Also try "Customer" as search
print("\n=== Finding entity type for 'Customer' ===")
raw2 = client.call_tool("data_find_entity_type", {"searchTerm": "Customer"})
if isinstance(raw2, str) and raw2.strip():
    result2 = json.loads(raw2)
    if isinstance(result2, list):
        for ent in result2[:8]:
            name = ent.get('Name', ent.get('name', ''))
            print(f"  {name}")
    else:
        print(f"  Result: {str(result2)[:500]}")

# Try "CustomersV3" which is a common D365 data entity
print("\n=== Getting metadata for 'CustomersV3' ===")
try:
    raw3 = client.call_tool("data_get_entity_metadata", {"entityName": "CustomersV3"})
    if isinstance(raw3, str) and raw3.strip():
        meta = json.loads(raw3)
        if isinstance(meta, dict):
            fields = meta.get("Properties", meta.get("properties", meta.get("Fields", [])))
            if isinstance(fields, list):
                # Search for Consignment/CHB fields
                for f in fields:
                    fname = f.get("Name", f.get("name", ""))
                    if any(t in fname.lower() for t in ["consignment", "chb", "Acme", "warehouse"]):
                        print(f"  FIELD: {fname}")
            elif isinstance(fields, dict):
                for fname in sorted(fields.keys()):
                    if any(t in fname.lower() for t in ["consignment", "chb", "Acme", "warehouse"]):
                        print(f"  FIELD: {fname}")
            print(f"  Total fields: {len(fields) if fields else 'N/A'}")
        else:
            print(f"  Meta type: {type(meta)}, content: {str(meta)[:300]}")
except Exception as e:
    print(f"  Error: {e}")

# Try querying a few customers to check
print("\n=== Querying customer 210205 in MY30 ===")
try:
    raw4 = client.call_tool("data_find_entities", {
        "entityName": "CustomersV3",
        "filter": "CustomerAccount eq '210205' and dataAreaId eq 'MY30'",
        "select": "CustomerAccount,CustomerGroupId,CHBConsignmentWarehouse,WarehouseId",
        "top": "1",
    })
    if isinstance(raw4, str) and raw4.strip():
        data = json.loads(raw4)
        print(f"  Result: {json.dumps(data, indent=2)[:1000]}")
except Exception as e:
    print(f"  Error: {e}")

# Also try for MY60 customer 1000
print("\n=== Querying customer 1000 in MY60 ===")
try:
    raw5 = client.call_tool("data_find_entities", {
        "entityName": "CustomersV3",
        "filter": "CustomerAccount eq '1000' and dataAreaId eq 'MY60'",
        "select": "CustomerAccount,CustomerGroupId,CHBConsignmentWarehouse,WarehouseId",
        "top": "1",
    })
    if isinstance(raw5, str) and raw5.strip():
        data = json.loads(raw5)
        print(f"  Result: {json.dumps(data, indent=2)[:1000]}")
except Exception as e:
    print(f"  Error: {e}")

print("\nDone.")
