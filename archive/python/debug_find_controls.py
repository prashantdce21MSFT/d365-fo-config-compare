"""Quick debug: test form_find_controls with various search terms."""
import sys, json
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient
from form_reader import open_form, close_form, mcp_call

with open(r"C:\D365DataValidator\config.json") as f:
    config = json.load(f)

client = D365McpClient(config["environments"]["ENV1"])
client.connect()

# Open Interest form
print("Opening Interest form...")
form_result = open_form(client, "Interest", "Display", "MY30")
fs = form_result.get("FormState", {})
print(f"Form: {fs.get('Name')} — {fs.get('Caption')}")

# Try various search terms
for term in ["*", "", "Interest", "CustInterest", "a", "Code"]:
    print(f"\n--- Search: '{term}' ---")
    try:
        raw = client.call_tool("form_find_controls", {"searchTerm": term})
        print(f"Raw type: {type(raw)}")
        print(f"Raw (first 500 chars): {str(raw)[:500]}")
        if isinstance(raw, str):
            try:
                parsed = json.loads(raw)
                print(f"Parsed keys: {list(parsed.keys()) if isinstance(parsed, dict) else 'not dict'}")
                if isinstance(parsed, dict):
                    for k, v in parsed.items():
                        if isinstance(v, list):
                            print(f"  {k}: {len(v)} items")
                            if v:
                                print(f"    First: {v[0]}")
                        elif isinstance(v, dict):
                            print(f"  {k}: {list(v.keys())[:10]}")
                        else:
                            print(f"  {k}: {v}")
            except:
                pass
    except Exception as e:
        print(f"Error: {e}")

close_form(client)
