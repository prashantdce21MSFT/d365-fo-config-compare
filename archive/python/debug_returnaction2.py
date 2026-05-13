"""Debug: ReturnActionDefaults controls."""
import sys, json
sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")
from d365_mcp_client import D365McpClient
from form_reader import open_form, close_form

with open(r"C:\D365DataValidator\config.json") as f:
    config = json.load(f)
client = D365McpClient(config["environments"]["ENV1"])
client.connect()

result = open_form(client, "purchReturnActionDefaults", "Display", "MY30")
fs = result.get("FormState", {})
print(f"Form: {fs.get('Name')} — {fs.get('Caption')}")
form = fs.get("Form", {})
print(f"Form structure (first 2000 chars): {json.dumps(form, indent=2)[:2000]}")

# Try search
for term in ["return", "action", "purch", "a", "e", "Default", "Disposition"]:
    raw = client.call_tool("form_find_controls", {"controlSearchTerm": term})
    if isinstance(raw, str):
        parsed = json.loads(raw) if raw.strip() else []
        if isinstance(parsed, list) and parsed:
            print(f"\n'{term}': {len(parsed)} controls")
            for c in parsed:
                print(f"  {c.get('Name')}")
        elif isinstance(parsed, dict) and "Result" not in parsed:
            print(f"\n'{term}': {parsed}")

close_form(client)
