# Step 4 — Configure form paths

Instead of remembering D365 menu item names (`mi=smmParameters`), paste UI navigation paths and let the extension discover the menu item via Playwright.

**Example paths**:

```
Production control > Setup > Production journal names
Production control > Setup > Production > Production pools
Sales and marketing > Setup > Sales and marketing parameters
```

When you click the button:

1. Pick an env to validate against
2. A scratch document opens — paste paths, one per line, save
3. Click **Validate** in the toast

Validated paths are saved to `.vscode/settings.json` (`d365FormExtractor.formPaths`) and offered as a QuickPick when you run **Extract Form**.

> Requires **Step 3 (Chrome CDP)** to be running.
