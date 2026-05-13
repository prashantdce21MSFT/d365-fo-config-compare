"""Debug: check ReturnAction form opening."""
import sys, json
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")
from d365_mcp_client import D365McpClient
from form_reader import open_form, close_form

with open(r"C:\D365DataValidator\config.json") as f:
    config = json.load(f)

client = D365McpClient(config["environments"]["ENV1"])
client.connect()

# Try different MI types
for mi_type in ["Display", "Action", "Output"]:
    print(f"\n--- Trying ReturnAction as {mi_type} ---")
    try:
        result = open_form(client, "ReturnAction", mi_type, "MY30")
        fs = result.get("FormState", {})
        print(f"  Name='{fs.get('Name')}', Caption='{fs.get('Caption')}'")
        form = fs.get("Form", {})
        print(f"  Form keys: {list(form.keys())[:10]}")
        # Try find_controls
        raw = client.call_tool("form_find_controls", {"controlSearchTerm": "return"})
        if isinstance(raw, str):
            parsed = json.loads(raw)
            if isinstance(parsed, list):
                print(f"  Controls found: {len(parsed)}")
                for c in parsed[:5]:
                    print(f"    {c.get('Name')}")
        close_form(client)
    except Exception as e:
        print(f"  Error: {e}")
        try: close_form(client)
        except: pass

# Also try the MI cache name
print("\n--- Checking MI cache ---")
mi_cache_file = r"C:\D365 Configuration Drift Analysis\output\batch_extract_mi_cache.json"
with open(mi_cache_file) as f:
    mi_cache = json.load(f)

# Find ReturnAction entries
for path, info in mi_cache.items():
    if "return" in path.lower() and "action" in path.lower():
        print(f"  Path: {path}")
        print(f"  Info: {info}")
