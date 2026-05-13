# Changelog

## 0.7.0 — 2026-05-13

- **Deep MCP extraction**: recursively walks every Tab / FastTab / Group / ReferenceGroup, merging nested fields and grids. Per-tab progress is logged.
- **Force-opens every tab** via `form_open_or_close_tab` (falls back to `form_click_control`) so previously-empty tabs now populate.
- **Full grid pagination** for every grid discovered at any nesting level (up to 200 pages, deduped).
- **Hidden-control sweep**: runs `form_find_controls` across ~146 alphanumeric + domain terms (Code, Name, Account, Vend, Cust, Item, …) to surface fields the tab walk missed. Output shows `Sweep added N new field(s)`.
- README rewrite and walkthrough polish.

## 0.6.x — 2026-05-13

- Multi-select form extraction in the Extract command.
- Form-path validator hardened: notification dismiss, magnifier-button fallback, breadcrumb-filtered result picker, MI-change detection, ASCII-safe output.
- Per-environment 4-tuple unpack fix.

## 0.5.1 — 2026-05-13

- Onboarding now actively drives `az login` per environment (tries default tenant first; prompts for tenant only on failure).
- Onboarding probes each environment's `/mcp` endpoint and records `mcpAvailable`.
- If no environment supports MCP, `defaultBackend` is auto-set to `playwright`.
- Onboarding offers inline form-path configuration + Playwright validation as a final step (no more typing menu items per extraction).
- Onboarding is idempotent: re-running it lets you keep existing environments/form paths or extend them.

## 0.1.0 — 2026-05-12

- Initial release.
- Two extraction backends: D365 ERP MCP (via Azure CLI auth) and Playwright (Chrome CDP).
- Onboarding wizard.
- Single-form extraction with optional cross-environment diff Excel.
- Start-Chrome-with-CDP command.
