"""Batch form comparison: Env1 UAT Asia (ENV1) vs Env4 Config (ENV4).

Reads pre-computed categorization from cat_checkpoint.json, compares
FORM items via MCP form tools. Checkpoints every 25 items.
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
CAT_FILE = r"C:\D365 Configuration Drift Analysis\cat_checkpoint.json"
OUTPUT_DIR = r"C:\D365 Configuration Drift Analysis\output"
CHECKPOINT_FILE = os.path.join(OUTPUT_DIR, "form_checkpoint.json")

SOURCE_ENV_KEY = "ENV1"
TARGET_ENV_KEY = "ENV4"
SOURCE_DISPLAY = "Env1 UAT Asia"
TARGET_DISPLAY = "Env4 Config"


# ── Company extraction ────────────────────────────────────────────────
_SKIP_TAGS = {"export", "import", "empties", "cutover"}
_COMPANY_MAP = {"MY": "MY30", "SG": "SG60"}

def extract_company_from_title(title):
    """Extract [MY], [SG], [MY60], [SG60], [MY30] etc. from title.
    Maps short codes to valid D365 legal entity IDs."""
    for match in re.finditer(r'\[([A-Za-z0-9]{2,6})\]', title):
        tag = match.group(1)
        if tag.lower() in _SKIP_TAGS:
            continue
        if tag.lower().startswith("fdd"):
            continue
        return _COMPANY_MAP.get(tag, tag)
    return None


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
    """Extract ADO IDs already processed from results keys."""
    ids = set()
    for key in results:
        r = results[key]
        if isinstance(r, dict) and r.get("item_id"):
            ids.add(str(r["item_id"]))
        elif key.startswith("FORM_"):
            ado = key[5:].split(" ")[0]
            ids.add(ado)
        elif key.startswith("Configuration Deliverable "):
            ado = key.split(":")[0].split()[-1]
            ids.add(ado)
    return ids


def compare_form_item(item, src_client, tgt_client, company_id="MY30"):
    """Compare a single form-based entity across two environments.

    Uses the pre-computed form_menu_item from categorization when available,
    otherwise falls back to resolve_menu_item_from_path.
    """
    title = item["title"]
    item_company = extract_company_from_title(title) or company_id

    # Try pre-computed menu item first
    menu_name = item.get("form_menu_item")
    menu_type = item.get("form_type") or "Display"
    paren_hint = None

    if not menu_name:
        # Fallback: resolve from title path
        menu_name, menu_type, paren_hint = resolve_menu_item_from_path(
            src_client, title, item_company
        )

    if not menu_name:
        return None, "NO_MENU_ITEM"

    # Read from source
    try:
        form_result = open_form(src_client, menu_name, menu_type or "Display", item_company)
        src_records, columns = read_form_records(
            src_client, form_result, include_details=True, tab_hint=paren_hint
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
        form_result = open_form(tgt_client, menu_name, menu_type or "Display", item_company)
        tgt_records, _ = read_form_records(
            tgt_client, form_result, include_details=True, tab_hint=paren_hint
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

    # Determine key fields from columns
    if columns:
        key_fields = columns[:2] if len(columns) > 2 else columns[:1]
    else:
        key_fields = []

    # Compare
    result = compare_entity_data(src_records, tgt_records, key_fields)
    return result, "OK"


# ── Main ──────────────────────────────────────────────────────────────
def main():
    print("=" * 70)
    print("  MY/SG BATCH FORM COMPARISON")
    print(f"  {SOURCE_DISPLAY} vs {TARGET_DISPLAY}")
    print("=" * 70)

    # 1. Load config
    with open(CONFIG_FILE) as f:
        config = json.load(f)

    # 2. Load categorization
    with open(CAT_FILE) as f:
        all_items = json.load(f)

    # Filter to FORM items (these have form_menu_item populated)
    form_items = [
        item for item in all_items
        if item.get("category") == "FORM"
        and item.get("form_menu_item")
    ]
    # Also try UNCERTAIN items that have a decent form_score but no odata match
    uncertain_form = [
        item for item in all_items
        if item.get("category") == "UNCERTAIN"
        and item.get("form_score", 0) >= 30
        and item.get("form_menu_item")
        and item.get("odata_score", 0) < 30  # not already tried in OData batch
    ]
    form_items.extend(uncertain_form)

    print(f"  Total categorized items: {len(all_items)}")
    print(f"  FORM items: {len(form_items) - len(uncertain_form)}")
    print(f"  UNCERTAIN (trying Form): {len(uncertain_form)}")
    print(f"  Total to process: {len(form_items)}")

    # 3. Connect to both environments
    print("\n  Connecting to environments...")
    src_client = D365McpClient(config["environments"][SOURCE_ENV_KEY])
    tgt_client = D365McpClient(config["environments"][TARGET_ENV_KEY])

    src_client.connect()
    print(f"  [OK] {SOURCE_DISPLAY}")

    tgt_client.connect()
    print(f"  [OK] {TARGET_DISPLAY}")

    # 4. Load checkpoint for resume
    results = load_checkpoint()
    already_done = done_ado_ids(results)
    skipped_count = 0

    # 5. Process each form item
    total = len(form_items)
    errors = []
    start_time = time.time()

    print(f"\n  {'='*70}")
    print(f"  Starting comparison ({total} items, {len(already_done)} already done)")
    print(f"  {'='*70}\n")

    for i, item in enumerate(form_items, 1):
        ado_id = str(item["id"])
        title = item["title"]

        if ado_id in already_done:
            skipped_count += 1
            continue

        company = extract_company_from_title(title) or "MY30"

        elapsed = time.time() - start_time
        processed = i - skipped_count
        if processed > 1:
            eta = (elapsed / (processed - 1)) * (total - i)
            eta_str = f"{eta/60:.1f}min"
        else:
            eta_str = "..."

        print(f"\n  [{i}/{total}] {ado_id}: {title[:80]}")
        print(f"    Menu: {item.get('form_menu_item', '?')}  Company: {company}  ETA: {eta_str}")

        try:
            result, status = compare_form_item(item, src_client, tgt_client, company)

            display_key = f"Configuration Deliverable {ado_id}: {title}"

            if result:
                result["item_id"] = ado_id
                result["category"] = item.get("category", "Form")
                results[display_key] = result

                s = result.get("summary", {})
                sc = s.get("source_count", 0)
                tc = s.get("target_count", 0)
                diffs = s.get("differences", 0) + s.get("only_in_source", 0) + s.get("only_in_target", 0)
                print(f"    >> {'DIFF' if diffs else 'MATCH'} (src:{sc} tgt:{tc} diffs:{diffs})")
            else:
                print(f"    >> {status}")
                if status not in ("NO_DATA", "SKIPPED"):
                    errors.append((ado_id, title_short, status))

        except Exception as e:
            print(f"    >> ERROR: {e}")
            errors.append((ado_id, title, str(e)))
            traceback.print_exc()

        # Checkpoint every 25 items
        processed_since_start = i - skipped_count
        if processed_since_start % 25 == 0 and processed_since_start > 0:
            save_checkpoint(results)
            elapsed = time.time() - start_time
            print(f"\n  --- Checkpoint saved ({len(results)} results, {elapsed:.0f}s elapsed) ---\n")

        time.sleep(0.3)  # gentle rate limit for form tools

    # 6. Final save + report
    save_checkpoint(results)

    print(f"\n  {'='*70}")
    print(f"  FORM COMPARISON COMPLETE")
    print(f"  {'='*70}")
    print(f"  Total form items: {total}")
    print(f"  Results collected: {len(results)}")
    print(f"  Errors: {len(errors)}")

    if errors:
        print(f"\n  Failed items:")
        for ado_id, title_text, err in errors[:30]:
            print(f"    {ado_id}: {title_text[:80]} -- {err[:80]}")

        error_log = os.path.join(OUTPUT_DIR, "form_errors.txt")
        with open(error_log, "w", encoding="utf-8") as f:
            f.write(f"Form Comparison Errors\n{'='*60}\n\n")
            for ado_id, title_text, err in errors:
                f.write(f"  {ado_id}: {title_text}\n    {err}\n\n")
        print(f"  Error log: {error_log}")

    if results:
        print(f"\n  Generating Excel report...")
        report_path = generate_comparison_report(
            results, SOURCE_DISPLAY, TARGET_DISPLAY, OUTPUT_DIR, report_tag="Form"
        )
        print(f"  Report saved: {report_path}")
    else:
        print("\n  No results to report.")

    elapsed = time.time() - start_time
    print(f"\n  Total time: {elapsed:.0f}s ({elapsed/60:.1f} min)")


if __name__ == "__main__":
    main()
