"""Test re-search for one DefaultDashboard item using last segment only, clicking via DOM."""

import re, sys
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

CDP_URL  = "http://localhost:9222"
HOME_URL = "https://example-uat.sandbox.operations.dynamics.com/?cmp=my30"
BASE_URL = "https://example-uat.sandbox.operations.dynamics.com/?cmp=my30&mi={mi}"

ADO_ID    = 36820
ADO_TITLE = "Warehouse management -> Setup -> Load posting methods"
SEARCH    = "Load posting methods"  # last segment only

def search_d365(page, search_term, ado_title):
    try:
        page.keyboard.press('Escape')
        page.wait_for_timeout(400)

        page.locator('button[aria-label="Search"]').click()
        page.wait_for_timeout(800)

        search_input = page.locator('input').filter(has_text='').first
        for sel in ['#NavigationSearchBox_searchBoxInput_input', 'input[id*="search" i]', 'input[placeholder*="earch" i]']:
            loc = page.locator(sel).first
            try:
                loc.wait_for(timeout=4000)
                search_input = loc
                break
            except:
                continue
        search_input.fill(search_term)
        page.wait_for_timeout(2500)

        # Score results and click via DOM (not pixel coords)
        ado_words = set(ado_title.lower().replace('>', ' ').split())
        result_data = page.evaluate("""() => {
            const results = [];
            document.querySelectorAll('li, [role="option"]').forEach((el, idx) => {
                const rect = el.getBoundingClientRect();
                if (rect.width === 0 || rect.height === 0) return;
                const txt = el.innerText.trim();
                if (txt && txt.length > 2 && txt.length < 200) {
                    results.push({text: txt, index: idx});
                }
            });
            return results;
        }""")

        print(f"Results found in dropdown: {len(result_data)}")
        for r in result_data[:10]:
            print(f"  [{r['index']}] {r['text'][:100]}")

        best = None
        best_score = -1
        for item in result_data:
            words = set(item['text'].lower().split())
            score = len(ado_words & words)
            if score > best_score:
                best_score = score
                best = item

        print(f"\nBest match (score={best_score}): {best['text'][:100] if best else 'None'}")

        if best and best_score > 0:
            # Click via DOM by index
            page.evaluate(f"""() => {{
                const all = Array.from(document.querySelectorAll('li, [role="option"]'));
                const el = all[{best['index']}];
                if (el) el.click();
            }}""")
        else:
            print("No good match, clicking first result position")
            page.mouse.click(992, 65)

        page.wait_for_load_state('domcontentloaded', timeout=15000)
        page.wait_for_timeout(2000)

        current_url = page.url
        print(f"\nFinal URL: {current_url}")
        mi_match = re.search(r'[?&]mi=([^&\s]+)', current_url)
        if mi_match:
            mi = mi_match.group(1)
            print(f"MI: {mi}")
            print(f"URL: {BASE_URL.format(mi=mi)}")
        else:
            print("No MI found in URL")

    except Exception as e:
        print(f"ERROR: {e}")

with sync_playwright() as pw:
    browser = pw.chromium.connect_over_cdp(CDP_URL)
    ctx  = browser.contexts[0]
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.set_viewport_size({'width': 1280, 'height': 720})
    page.bring_to_front()
    page.goto(HOME_URL, wait_until='domcontentloaded', timeout=30000)
    page.wait_for_timeout(2000)
    if 'login.microsoftonline' in page.url:
        print("ERROR: Not signed in")
        sys.exit(1)
    print(f"Searching for: '{SEARCH}' (ADO: {ADO_TITLE})\n")
    search_d365(page, SEARCH, ADO_TITLE)
    browser.close()
