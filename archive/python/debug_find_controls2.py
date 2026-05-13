"""Quick debug: test form_find_controls with correct parameter name."""
import sys, json
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient
from form_reader import open_form, close_form

with open(r"C:\D365DataValidator\config.json") as f:
    config = json.load(f)

client = D365McpClient(config["environments"]["ENV1"])
client.connect()

# Open Interest form
print("Opening Interest form...")
form_result = open_form(client, "Interest", "Display", "MY30")
fs = form_result.get("FormState", {})
print(f"Form: {fs.get('Name')} — {fs.get('Caption')}")

# Try with correct param name
for term in ["*", "a", "Interest", "Code"]:
    print(f"\n--- Search: '{term}' ---")
    try:
        raw = client.call_tool("form_find_controls", {"controlSearchTerm": term})
        if isinstance(raw, str):
            parsed = json.loads(raw)
            # Look for control data
            for k, v in parsed.items():
                if k in ("SessionId", "MCPActivityId"):
                    continue
                if isinstance(v, list):
                    print(f"  {k}: {len(v)} items")
                    for item in v[:3]:
                        print(f"    {item}")
                elif isinstance(v, dict):
                    subkeys = list(v.keys())[:10]
                    print(f"  {k}: keys={subkeys}")
                else:
                    print(f"  {k}: {str(v)[:200]}")
    except Exception as e:
        print(f"Error: {e}")

client.call_tool("form_close_form", {})
