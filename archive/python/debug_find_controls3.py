"""Debug: inspect form_find_controls response structure."""
import sys, json
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient
from form_reader import open_form, close_form

with open(r"C:\D365DataValidator\config.json") as f:
    config = json.load(f)

client = D365McpClient(config["environments"]["ENV1"])
client.connect()

print("Opening Interest form...")
form_result = open_form(client, "Interest", "Display", "MY30")
fs = form_result.get("FormState", {})
print(f"Form: {fs.get('Name')} — {fs.get('Caption')}")

# Get raw response
for term in ["a", "Interest", "Code", "Cust"]:
    print(f"\n--- Search: '{term}' ---")
    try:
        raw = client.call_tool("form_find_controls", {"controlSearchTerm": term})
        print(f"Raw type: {type(raw)}")
        if isinstance(raw, str):
            print(f"Raw (first 800): {raw[:800]}")
            try:
                parsed = json.loads(raw)
                print(f"Parsed type: {type(parsed)}")
                if isinstance(parsed, list):
                    print(f"List length: {len(parsed)}")
                    for item in parsed[:5]:
                        print(f"  Item: {item}")
                elif isinstance(parsed, dict):
                    print(f"Dict keys: {list(parsed.keys())}")
            except:
                pass
        else:
            print(f"Not string: {str(raw)[:800]}")
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

client.call_tool("form_close_form", {})
