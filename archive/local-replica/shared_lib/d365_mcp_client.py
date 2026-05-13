"""
D365 FO MCP Client — handles OAuth2 authentication and MCP session management.
"""
import time
import json
import requests


class D365McpClient:
    """Client for interacting with Dynamics 365 F&O MCP Server."""

    def __init__(self, env_config):
        self.mcp_url = env_config["mcp_url"]
        self.resource_url = env_config["resource_url"]
        self.tenant_id = env_config["tenant_id"]
        self.client_id = env_config["client_id"]
        self.client_secret = env_config["client_secret"]
        self.env_name = env_config.get("name", "Unknown")

        self._access_token = None
        self._token_expiry = 0
        self._session_id = None

    # ── OAuth2 ──────────────────────────────────────────────────────────

    def _get_token(self):
        """Acquire or reuse an OAuth2 access token via client credentials."""
        if self._access_token and time.time() < self._token_expiry - 60:
            return self._access_token

        token_url = (
            f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        )
        payload = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": f"{self.resource_url}/.default",
        }
        resp = requests.post(token_url, data=payload, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        self._access_token = data["access_token"]
        self._token_expiry = time.time() + data.get("expires_in", 3600)
        return self._access_token

    # ── MCP Session ─────────────────────────────────────────────────────

    def _mcp_request(self, method, params=None, rpc_id=None):
        """Send a JSON-RPC request to the MCP server."""
        token = self._get_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
        if self._session_id:
            headers["mcp-session-id"] = self._session_id

        body = {"jsonrpc": "2.0", "method": method}
        if params is not None:
            body["params"] = params
        if rpc_id is not None:
            body["id"] = rpc_id

        resp = requests.post(self.mcp_url, json=body, headers=headers, timeout=120)
        resp.raise_for_status()

        # Capture session id from headers
        sid = resp.headers.get("mcp-session-id")
        if sid:
            self._session_id = sid

        if rpc_id is not None:
            return resp.json()
        return None  # notifications have no response body to parse

    def connect(self):
        """Initialize the MCP session."""
        result = self._mcp_request(
            "initialize",
            params={
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "D365DataValidator", "version": "1.0"},
            },
            rpc_id=1,
        )
        # Send initialized notification
        self._mcp_request("notifications/initialized", params={})
        server_info = result.get("result", {}).get("serverInfo", {})
        print(
            f"  Connected to {server_info.get('name', 'unknown')} "
            f"v{server_info.get('version', '?')} [{self.env_name}]"
        )
        return result

    # ── Tool Calls ──────────────────────────────────────────────────────

    def call_tool(self, tool_name, arguments, rpc_id=2):
        """Call an MCP tool and return the result content."""
        resp = self._mcp_request(
            "tools/call",
            params={"name": tool_name, "arguments": arguments},
            rpc_id=rpc_id,
        )
        if "error" in resp:
            raise RuntimeError(
                f"MCP tool error ({tool_name}): {resp['error'].get('message', resp['error'])}"
            )
        # Extract text content from MCP tool result
        content_list = resp.get("result", {}).get("content", [])
        texts = [c.get("text", "") for c in content_list if c.get("type") == "text"]
        return "\n".join(texts)

    def find_entity_types(self, search_filter, top=10):
        """Search for OData entity types."""
        raw = self.call_tool(
            "data_find_entity_type",
            {"tableSearchFilter": search_filter, "topHitCount": str(top)},
        )
        return raw

    def get_entity_metadata(self, entity_set_name, include_keys=True):
        """Get metadata (fields, keys) for an entity."""
        raw = self.call_tool(
            "data_get_entity_metadata",
            {
                "entitySetName": entity_set_name,
                "includeKeys": str(include_keys).lower(),
                "includeFieldConstraints": "false",
                "includeRelationships": "false",
                "includeEnumValues": "false",
            },
        )
        return raw

    def find_entities(self, odata_path, query_options=None):
        """Query entity data via OData."""
        args = {"odataPath": odata_path}
        if query_options:
            args["odataQueryOptions"] = query_options
        raw = self.call_tool("data_find_entities", args)
        return raw
