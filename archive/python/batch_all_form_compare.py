"""Batch form comparison for ALL 606 MY/SG Configuration Deliverables.

Uses MCP form tools (not OData) for every item:
1. Resolves menu item from ADO title navigation path
2. Opens form on both UAT (ENV1) and Config (ENV4)
3. Reads all records with pagination (>25 rows)
4. Compares field-by-field
5. Checkpoints every 10 items
"""
import sys
import os
import json
import re
import time
import traceback

sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient
from comparator import compare_entity_data
from form_reader import (
    open_form, close_form, read_form_records,
    resolve_menu_item_from_path, find_menu_item,
)
from mysg_reporter import generate_report as generate_comparison_report

# ── Config ────────────────────────────────────────────────────────────
CONFIG_FILE = r"C:\D365DataValidator\config.json"
ADO_ITEMS_FILE = r"C:\D365 Configuration Drift Analysis\output\all_ado_items.json"
OUTPUT_DIR = r"C:\D365 Configuration Drift Analysis\output"
CHECKPOINT_FILE = os.path.join(OUTPUT_DIR, "all_form_checkpoint.json")
MI_CACHE_FILE = os.path.join(OUTPUT_DIR, "mi_cache.json")

SOURCE_ENV_KEY = "ENV1"
TARGET_ENV_KEY = "ENV4"
SOURCE_DISPLAY = "Env1 UAT Asia"
TARGET_DISPLAY = "Env4 Config"


# ── Company extraction ────────────────────────────────────────────────
_SKIP_TAGS = {"export", "import", "empties", "cutover"}
_COMPANY_MAP = {"MY": "MY30", "SG": "SG60"}
# Additional title patterns for company (not in brackets)
_COMPANY_PATTERNS = [
    (r'\bMY30\b', 'MY30'), (r'\bSG60\b', 'SG60'),
    (r'\bMY60\b', 'MY60'), (r'\bMY\d{2}\b', None),
    (r'\(MY60\)', 'MY60'), (r'\(SG60\)', 'SG60'),
    (r'\(MY30\)', 'MY30'),
]

def extract_company(title):
    """Extract company from title — brackets first, then patterns."""
    for match in re.finditer(r'\[([A-Za-z0-9]{2,6})\]', title):
        tag = match.group(1)
        if tag.lower() in _SKIP_TAGS:
            continue
        if tag.lower().startswith("fdd"):
            continue
        return _COMPANY_MAP.get(tag, tag)
    # Check for (MY60), (SG60) etc in parentheses
    for pattern, company in _COMPANY_PATTERNS:
        if company and re.search(pattern, title):
            return company
    # Check title text for company mentions
    title_lower = title.lower()
    if 'sg60' in title_lower or '(sg60)' in title_lower:
        return 'SG60'
    if 'my30' in title_lower or '(my30)' in title_lower:
        return 'MY30'
    if 'my60' in title_lower or '(my60)' in title_lower:
        return 'MY60'
    return None


# ── MI cache (avoid re-resolving menu items) ──────────────────────────
def load_mi_cache():
    if os.path.isfile(MI_CACHE_FILE):
        try:
            with open(MI_CACHE_FILE, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_mi_cache(cache):
    with open(MI_CACHE_FILE, "w") as f:
        json.dump(cache, f, indent=2)


# ── Checkpoint helpers ────────────────────────────────────────────────
def load_checkpoint():
    if os.path.isfile(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r") as f:
                data = json.load(f)
            print(f"  Loaded checkpoint: {len(data)} completed items")
            return data
        except Exception:
            pass
    return {}


def save_checkpoint(results):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(CHECKPOINT_FILE, "w") as f:
        json.dump(results, f, default=str)


def done_ado_ids(results):
    """Extract ADO IDs already processed."""
    ids = set()
    for key, r in results.items():
        if isinstance(r, dict) and r.get("item_id"):
            ids.add(str(r["item_id"]))
        elif isinstance(r, dict) and r.get("ado_id"):
            ids.add(str(r["ado_id"]))
    # Also check skip/error entries
    for key in results:
        if key.startswith("Configuration Deliverable "):
            ado = key.split(":")[0].split()[-1]
            ids.add(ado)
    return ids


# ── Core comparison ──────────────────────────────────────────────────
def compare_form_item(title, menu_name, menu_type, paren_hint, company,
                      src_client, tgt_client):
    """Compare a form across two environments. Returns (result_dict, status)."""

    # Read from source
    try:
        form_result = open_form(src_client, menu_name, menu_type or "Display", company)
        src_records, columns = read_form_records(
            src_client, form_result, include_details=False, tab_hint=paren_hint
        )
        close_form(src_client)
    except Exception as e:
        try:
            close_form(src_client)
        except Exception:
            pass
        return None, f"SRC_ERROR: {e}"

    print(f"    source: {len(src_records)} records")

    # Read from target
    try:
        form_result = open_form(tgt_client, menu_name, menu_type or "Display", company)
        tgt_records, _ = read_form_records(
            tgt_client, form_result, include_details=False, tab_hint=paren_hint
        )
        close_form(tgt_client)
        print(f"    target: {len(tgt_records)} records")
    except Exception as e:
        try:
            close_form(tgt_client)
        except Exception:
            pass
        print(f"    target: ERROR - {e}")
        tgt_records = []

    if not src_records and not tgt_records:
        return None, "NO_DATA"

    # Use columns as key fields (first 2 columns for grid forms)
    if columns:
        key_fields = columns[:2] if len(columns) > 2 else columns[:1]
    else:
        key_fields = []

    result = compare_entity_data(src_records, tgt_records, key_fields)
    return result, "OK"


# ── Main ──────────────────────────────────────────────────────────────
def main():
    print("=" * 70)
    print("  MY/SG BATCH FORM COMPARISON — ALL 606 ITEMS")
    print(f"  {SOURCE_DISPLAY} vs {TARGET_DISPLAY}")
    print("=" * 70)

    # 1. Load config
    with open(CONFIG_FILE) as f:
        config = json.load(f)

    # 2. Load all ADO items
    with open(ADO_ITEMS_FILE) as f:
        all_items = json.load(f)
    print(f"  Total ADO items: {len(all_items)}")

    # 3. Connect to both environments
    print("\n  Connecting to environments...")
    src_client = D365McpClient(config["environments"][SOURCE_ENV_KEY])
    tgt_client = D365McpClient(config["environments"][TARGET_ENV_KEY])

    src_client.connect()
    print(f"  [OK] {SOURCE_DISPLAY}")

    tgt_client.connect()
    print(f"  [OK] {TARGET_DISPLAY}")

    # 4. Load checkpoint + MI cache
    results = load_checkpoint()
    already_done = done_ado_ids(results)
    mi_cache = load_mi_cache()

    total = len(all_items)
    errors = []
    no_menu = []
    no_data = []
    start_time = time.time()
    processed_count = 0

    print(f"\n  {'='*70}")
    print(f"  Starting comparison ({total} items, {len(already_done)} already done)")
    print(f"  {'='*70}\n")

    for i, item in enumerate(all_items, 1):
        ado_id = str(item["ado_id"])
        title = item["title"]

        if ado_id in already_done:
            continue

        processed_count += 1
        company = extract_company(title) or "MY30"

        # ETA
        elapsed = time.time() - start_time
        if processed_count > 1:
            eta = (elapsed / (processed_count - 1)) * (total - i)
            eta_str = f"{eta/60:.1f}min"
        else:
            eta_str = "..."

        print(f"\n  [{i}/{total}] {ado_id}: {title[:80]}")

        # Resolve menu item (check cache first)
        cache_key = ado_id
        if cache_key in mi_cache:
            mi_info = mi_cache[cache_key]
            menu_name = mi_info.get("menu_name")
            menu_type = mi_info.get("menu_type", "Display")
            paren_hint = mi_info.get("paren_hint")
            if not menu_name:
                print(f"    >> SKIP (no menu item, cached)")
                no_menu.append((ado_id, title))
                continue
        else:
            # Use nav_path from title
            nav_path = item.get("nav_path", "")
            if not nav_path:
                # No path in title — skip for now (needs screenshot analysis)
                print(f"    >> SKIP (no navigation path in title)")
                mi_cache[cache_key] = {"menu_name": None, "reason": "no_nav_path"}
                no_menu.append((ado_id, title))
                continue

            print(f"    Resolving menu item from: {nav_path[:70]}...")
            try:
                menu_name, menu_type, paren_hint = resolve_menu_item_from_path(
                    src_client, nav_path, company
                )
            except Exception as e:
                print(f"    >> ERROR resolving: {e}")
                mi_cache[cache_key] = {"menu_name": None, "reason": f"error: {e}"}
                errors.append((ado_id, title, f"RESOLVE_ERROR: {e}"))
                continue

            # Cache the result
            mi_cache[cache_key] = {
                "menu_name": menu_name,
                "menu_type": menu_type or "Display",
                "paren_hint": paren_hint,
            }

            if not menu_name:
                print(f"    >> SKIP (could not resolve menu item)")
                no_menu.append((ado_id, title))
                continue

        print(f"    Menu: {menu_name}  Company: {company}  ETA: {eta_str}")

        try:
            result, status = compare_form_item(
                title, menu_name, menu_type, paren_hint, company,
                src_client, tgt_client
            )

            display_key = f"Configuration Deliverable {ado_id}: {title}"

            if result:
                result["item_id"] = ado_id
                result["menu_item"] = menu_name
                result["company"] = company
                results[display_key] = result

                s = result.get("summary", {})
                sc = s.get("source_count", 0)
                tc = s.get("target_count", 0)
                diffs = s.get("differences", 0) + s.get("only_in_source", 0) + s.get("only_in_target", 0)
                print(f"    >> {'DIFF' if diffs else 'MATCH'} (src:{sc} tgt:{tc} diffs:{diffs})")
            else:
                print(f"    >> {status}")
                if status == "NO_DATA":
                    no_data.append((ado_id, title))
                    # Still record it so we don't retry
                    results[display_key] = {
                        "item_id": ado_id, "menu_item": menu_name,
                        "company": company, "status": "NO_DATA",
                        "summary": {"source_count": 0, "target_count": 0,
                                    "differences": 0, "only_in_source": 0,
                                    "only_in_target": 0},
                    }
                elif "ERROR" in status:
                    errors.append((ado_id, title, status))

        except Exception as e:
            print(f"    >> ERROR: {e}")
            errors.append((ado_id, title, str(e)))
            traceback.print_exc()

        # Checkpoint every 5 items
        if processed_count % 5 == 0:
            save_checkpoint(results)
            save_mi_cache(mi_cache)
            elapsed = time.time() - start_time
            done_pct = len(done_ado_ids(results)) / total * 100
            print(f"\n  --- Checkpoint ({len(results)} results, {done_pct:.0f}%, {elapsed:.0f}s) ---\n")

        time.sleep(0.3)

    # 6. Final save
    save_checkpoint(results)
    save_mi_cache(mi_cache)

    # Stats
    all_done = done_ado_ids(results)
    print(f"\n  {'='*70}")
    print(f"  FORM COMPARISON COMPLETE")
    print(f"  {'='*70}")
    print(f"  Total items: {total}")
    print(f"  Completed: {len(all_done)}")
    print(f"  With results: {sum(1 for r in results.values() if isinstance(r, dict) and r.get('summary', {}).get('source_count', 0) > 0)}")
    print(f"  No data: {len(no_data)}")
    print(f"  No menu item: {len(no_menu)}")
    print(f"  Errors: {len(errors)}")

    if no_menu:
        print(f"\n  Items without menu item ({len(no_menu)}):")
        for ado_id, title_text in no_menu[:20]:
            print(f"    {ado_id}: {title_text[:80]}")
        if len(no_menu) > 20:
            print(f"    ... and {len(no_menu)-20} more")

    if errors:
        print(f"\n  Errors ({len(errors)}):")
        for ado_id, title_text, err in errors[:20]:
            print(f"    {ado_id}: {title_text[:60]} -- {err[:60]}")

        error_log = os.path.join(OUTPUT_DIR, "all_form_errors.txt")
        with open(error_log, "w", encoding="utf-8") as f:
            f.write(f"All Form Comparison Errors\n{'='*60}\n\n")
            for ado_id, title_text, err in errors:
                f.write(f"  {ado_id}: {title_text}\n    {err}\n\n")

    # Generate report
    report_results = {k: v for k, v in results.items()
                      if isinstance(v, dict) and v.get("summary", {}).get("source_count", 0) > 0}
    if report_results:
        print(f"\n  Generating Excel report ({len(report_results)} items with data)...")
        report_path = generate_comparison_report(
            report_results, SOURCE_DISPLAY, TARGET_DISPLAY, OUTPUT_DIR,
            report_tag="AllForm"
        )
        print(f"  Report saved: {report_path}")

    elapsed = time.time() - start_time
    print(f"\n  Total time: {elapsed:.0f}s ({elapsed/60:.1f} min)")


if __name__ == "__main__":
    main()
