# Step 3 — Launch Chrome with remote debugging

The **Playwright backend** drives a Chrome instance running on `localhost:9222`.

This step:

1. Closes any open Chrome windows
2. Relaunches Chrome with `--remote-debugging-port=9222`
3. Uses a dedicated profile (`<workspace>/_chrome_profile`)

After Chrome opens, **sign in to your D365 environment** in that window. Your session stays alive for subsequent extractions.

> Skip this step if you only plan to use the **D365 ERP MCP** backend.
