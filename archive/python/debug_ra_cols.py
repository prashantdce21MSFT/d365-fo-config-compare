"""Quick check: what are the 2 control names for ReturnAction?"""
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
form = result.get("FormState", {}).get("Form", {})
# Get grid columns
grid = form.get("Grid", {})
for gname, ginfo in grid.items():
    if isinstance(ginfo, dict):
        for col in ginfo.get("Columns", []):
            print(f"Column: {col.get('Name')} — {col.get('Label')}")
close_form(client)
