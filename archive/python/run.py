"""
MY/SG D365 Configuration Drift Analysis — Main Orchestrator.

Compares configuration across Env1 UAT Asia (ENV1) vs Env4 Config (ENV4)
using pre-computed categorization from cat_checkpoint.json.

Usage:
  python run.py              # Run all (OData first, then Form)
  python run.py odata        # OData items only
  python run.py form         # Form items only
"""
import sys
import os
import json
import re
import time
import traceback
from datetime import datetime

sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from d365_mcp_client import D365McpClient
from data_fetcher import fetch_entity_data, select_safe_fields, print_progress_bar
from comparator import compare_entity_data
from form_reader import (
    open_form, close_form, read_form_records,
    resolve_menu_item_from_path,
)
from mysg_reporter import generate_report as generate_comparison_report

# ── Config ────────────────────────────────────────────────────────────
CONFIG_FILE = r"C:\D365DataValidator\config.json"
CAT_FILE = r"C:\D365 Configuration Drift Analysis\cat_checkpoint.json"
OUTPUT_DIR = r"C:\D365 Configuration Drift Analysis\output"

SOURCE_ENV_KEY = "ENV1"
TARGET_ENV_KEY = "ENV4"
SOURCE_DISPLAY = "Env1 UAT Asia"
TARGET_DISPLAY = "Env4 Config"

_SKIP_TAGS = {"export", "import", "empties", "cutover"}
_COMPANY_MAP = {"MY": "MY30", "SG": "SG60"}


def extract_company_from_title(title):
    """Extract [MY], [SG], [MY60], [SG60], [MY30] etc. from title.
    Maps short codes to valid D365 legal entity IDs."""
    for match in re.finditer(r'\[([A-Za-z0-9]{2,6})\]', title):
        tag = match.group(1)
        if tag.lower() in _SKIP_TAGS or tag.lower().startswith("fdd"):
            continue
        return _COMPANY_MAP.get(tag, tag)
    return None


def format_duration(seconds):
    if seconds < 60:
        return f"{seconds:.0f}s"
    return f"{int(seconds//60)}m {int(seconds%60)}s"


# ── Checkpoint ────────────────────────────────────────────────────────
def load_checkpoint(path):
    if os.path.isfile(path):
        try:
            with open(path, "r") as f:
                data = json.load(f)
            return data
        except Exception:
            pass
    return {}


def save_checkpoint(results, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w") as f:
        json.dump(results, f, default=str)


def done_ado_ids(results, prefix=""):
    ids = set()
    for key in results:
        r = results[key]
        if isinstance(r, dict) and r.get("item_id"):
            ids.add(str(r["item_id"]))
        elif prefix and key.startswith(prefix):
            ado = key[len(prefix):].split(" ")[0]
            ids.add(ado)
        elif key.startswith("Configuration Deliverable "):
            ado = key.split(":")[0].split()[-1]
            ids.add(ado)
    return ids


def extract_key_fields(client, entity_name):
    raw = client.get_entity_metadata(entity_name, include_keys=True)
    selected, keys = select_safe_fields(raw, max_fields=25)
    return {"keys": keys, "select_fields": selected}


# ── OData comparison ──────────────────────────────────────────────────
def run_odata_phase(odata_items, src_client, tgt_client, config):
    checkpoint_file = os.path.join(OUTPUT_DIR, "odata_checkpoint.json")
    results = load_checkpoint(checkpoint_file)
    already_done = done_ado_ids(results, "ODATA_")
    max_records = config.get("settings", {}).get("max_records_per_entity", 50000)

    total = len(odata_items)
    errors = []
    skipped = 0
    start_time = time.time()

    print(f"\n{'='*70}")
    print(f"  PHASE: OData ({total} items, {len(already_done)} already done)")
    print(f"{'='*70}")

    for i, item in enumerate(odata_items, 1):
        ado_id = str(item["id"])
        title = item["title"]
        entity_name = item["odata_entity"]

        if ado_id in already_done:
            skipped += 1
            continue

        company = extract_company_from_title(title)
        le_filter = f"dataAreaId eq '{company.lower()}'" if company else None

        processed = i - skipped
        if processed > 1:
            elapsed = time.time() - start_time
            eta = format_duration((elapsed / (processed - 1)) * (total - i))
        else:
            eta = "..."

        print(f"\n  [{i}/{total}] {ado_id}: {title[:80]}")
        print(f"    Entity: {entity_name}  Company: {company or 'cross-company'}  ETA: {eta}")

        try:
            fields_info = extract_key_fields(src_client, entity_name)

            sys.stdout.write(f"    source: ")
            src_records = fetch_entity_data(
                src_client, entity_name,
                select_fields=fields_info.get("select_fields"),
                filter_expr=le_filter, max_records=max_records,
                progress_callback=print_progress_bar,
            )

            sys.stdout.write(f"    target: ")
            tgt_records = fetch_entity_data(
                tgt_client, entity_name,
                select_fields=fields_info.get("select_fields"),
                filter_expr=le_filter, max_records=max_records,
                progress_callback=print_progress_bar,
            )

            if not src_records and not tgt_records:
                print(f"    >> No data")
                continue

            result = compare_entity_data(src_records, tgt_records, fields_info["keys"])
            display_key = f"Configuration Deliverable {ado_id}: {title}"
            result["item_id"] = ado_id
            result["category"] = item.get("category", "OData")
            results[display_key] = result

            s = result.get("summary", {})
            diffs = s.get("differences", 0) + s.get("only_in_source", 0) + s.get("only_in_target", 0)
            print(f"    >> {'DIFF' if diffs else 'MATCH'} (src:{s.get('source_count',0)} tgt:{s.get('target_count',0)} diffs:{diffs})")

        except Exception as e:
            print(f"    >> ERROR: {e}")
            errors.append((ado_id, title, str(e)))

        if (i - skipped) % 25 == 0 and (i - skipped) > 0:
            save_checkpoint(results, checkpoint_file)
            print(f"\n  --- Checkpoint ({len(results)} results) ---\n")

    save_checkpoint(results, checkpoint_file)

    # Generate report
    report_path = None
    if results:
        print(f"\n  Generating OData report...")
        report_path = generate_comparison_report(
            results, SOURCE_DISPLAY, TARGET_DISPLAY, OUTPUT_DIR, report_tag="OData"
        )

    phase_time = time.time() - start_time
    print(f"\n  OData phase: {len(results)} results, {len(errors)} errors in {format_duration(phase_time)}")
    return results, errors, report_path


# ── Form comparison ───────────────────────────────────────────────────
def run_form_phase(form_items, src_client, tgt_client):
    checkpoint_file = os.path.join(OUTPUT_DIR, "form_checkpoint.json")
    results = load_checkpoint(checkpoint_file)
    already_done = done_ado_ids(results, "FORM_")

    total = len(form_items)
    errors = []
    skipped = 0
    start_time = time.time()

    print(f"\n{'='*70}")
    print(f"  PHASE: Form ({total} items, {len(already_done)} already done)")
    print(f"{'='*70}")

    for i, item in enumerate(form_items, 1):
        ado_id = str(item["id"])
        title = item["title"]

        if ado_id in already_done:
            skipped += 1
            continue

        company = extract_company_from_title(title) or "MY30"

        processed = i - skipped
        if processed > 1:
            elapsed = time.time() - start_time
            eta = format_duration((elapsed / (processed - 1)) * (total - i))
        else:
            eta = "..."

        print(f"\n  [{i}/{total}] {ado_id}: {title[:80]}")
        print(f"    Menu: {item.get('form_menu_item', '?')}  Company: {company}  ETA: {eta}")

        try:
            menu_name = item.get("form_menu_item")
            menu_type = item.get("form_type") or "Display"
            paren_hint = None

            if not menu_name:
                menu_name, menu_type, paren_hint = resolve_menu_item_from_path(
                    src_client, title, company
                )

            if not menu_name:
                print(f"    >> NO_MENU_ITEM")
                errors.append((ado_id, title, "Menu item not found"))
                continue

            # Source
            form_result = open_form(src_client, menu_name, menu_type or "Display", company)
            src_records, columns = read_form_records(src_client, form_result, include_details=True, tab_hint=paren_hint)
            close_form(src_client)
            print(f"    source: {len(src_records)} records")

            # Target
            try:
                form_result = open_form(tgt_client, menu_name, menu_type or "Display", company)
                tgt_records, _ = read_form_records(tgt_client, form_result, include_details=True, tab_hint=paren_hint)
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
                print(f"    >> No data")
                continue

            key_fields = columns[:2] if columns and len(columns) > 2 else (columns[:1] if columns else [])
            result = compare_entity_data(src_records, tgt_records, key_fields)

            display_key = f"Configuration Deliverable {ado_id}: {title}"
            result["item_id"] = ado_id
            result["category"] = item.get("category", "Form")
            results[display_key] = result

            s = result.get("summary", {})
            diffs = s.get("differences", 0) + s.get("only_in_source", 0) + s.get("only_in_target", 0)
            print(f"    >> {'DIFF' if diffs else 'MATCH'} (src:{s.get('source_count',0)} tgt:{s.get('target_count',0)} diffs:{diffs})")

        except Exception as e:
            print(f"    >> ERROR: {e}")
            errors.append((ado_id, title, str(e)))
            try:
                close_form(src_client)
            except Exception:
                pass

        if (i - skipped) % 25 == 0 and (i - skipped) > 0:
            save_checkpoint(results, checkpoint_file)
            print(f"\n  --- Checkpoint ({len(results)} results) ---\n")

        time.sleep(0.3)

    save_checkpoint(results, checkpoint_file)

    report_path = None
    if results:
        print(f"\n  Generating Form report...")
        report_path = generate_comparison_report(
            results, SOURCE_DISPLAY, TARGET_DISPLAY, OUTPUT_DIR, report_tag="Form"
        )

    phase_time = time.time() - start_time
    print(f"\n  Form phase: {len(results)} results, {len(errors)} errors in {format_duration(phase_time)}")
    return results, errors, report_path


# ── Main ──────────────────────────────────────────────────────────────
def main():
    mode = sys.argv[1].lower() if len(sys.argv) > 1 else "all"

    print("=" * 70)
    print("  MY/SG D365 CONFIGURATION DRIFT ANALYSIS")
    print(f"  {SOURCE_DISPLAY} vs {TARGET_DISPLAY}")
    print(f"  Mode: {mode.upper()}")
    print("=" * 70)

    # Load config + categorization
    with open(CONFIG_FILE) as f:
        config = json.load(f)
    with open(CAT_FILE) as f:
        all_items = json.load(f)

    # Split items by category
    odata_items = [i for i in all_items if i.get("category") in ("BOTH", "ODATA") and i.get("odata_entity")]
    form_items = [i for i in all_items if i.get("category") == "FORM" and i.get("form_menu_item")]
    uncertain = [i for i in all_items if i.get("category") == "UNCERTAIN"]
    not_found = [i for i in all_items if i.get("category") == "NOT_FOUND"]

    # Route UNCERTAIN items
    uncertain_odata = [i for i in uncertain if i.get("odata_score", 0) >= 30 and i.get("odata_entity")]
    uncertain_form = [i for i in uncertain if i.get("form_score", 0) >= 30 and i.get("form_menu_item") and i.get("odata_score", 0) < 30]
    odata_items.extend(uncertain_odata)
    form_items.extend(uncertain_form)

    print(f"\n  Categorization breakdown:")
    print(f"    OData (BOTH+ODATA+uncertain): {len(odata_items)}")
    print(f"    Form (FORM+uncertain):        {len(form_items)}")
    print(f"    NOT_FOUND (skipped):          {len(not_found)}")
    print(f"    Unrouted UNCERTAIN:           {len(uncertain) - len(uncertain_odata) - len(uncertain_form)}")

    # Connect
    print(f"\n  Connecting to environments...")
    src_client = D365McpClient(config["environments"][SOURCE_ENV_KEY])
    tgt_client = D365McpClient(config["environments"][TARGET_ENV_KEY])
    src_client.connect()
    print(f"  [OK] {SOURCE_DISPLAY}")
    tgt_client.connect()
    print(f"  [OK] {TARGET_DISPLAY}")

    start_time = time.time()
    all_errors = []
    report_paths = []

    # Run phases
    if mode in ("all", "odata") and odata_items:
        odata_results, odata_errors, odata_report = run_odata_phase(odata_items, src_client, tgt_client, config)
        all_errors.extend(odata_errors)
        if odata_report:
            report_paths.append(odata_report)

    if mode in ("all", "form") and form_items:
        form_results, form_errors, form_report = run_form_phase(form_items, src_client, tgt_client)
        all_errors.extend(form_errors)
        if form_report:
            report_paths.append(form_report)

    # Summary
    total_time = time.time() - start_time
    print(f"\n{'='*70}")
    print(f"  COMPLETE — {format_duration(total_time)}")
    print(f"{'='*70}")
    print(f"  Errors: {len(all_errors)}")
    for rp in report_paths:
        print(f"  Report: {rp}")

    if all_errors:
        print(f"\n  Errors ({len(all_errors)}):")
        for ado_id, title_text, err in all_errors[:30]:
            print(f"    {ado_id}: {title_text[:80]} -- {err[:80]}")

        error_log = os.path.join(OUTPUT_DIR, f"error_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        with open(error_log, "w", encoding="utf-8") as f:
            f.write(f"MY/SG Config Drift — Error Log\n{'='*60}\n\n")
            for ado_id, title_text, err in all_errors:
                f.write(f"  {ado_id}: {title_text}\n    {err}\n\n")
        print(f"  Error log: {error_log}")

    if not_found:
        print(f"\n  NOT_FOUND items ({len(not_found)}) — manual review needed:")
        for item in not_found[:10]:
            print(f"    {item['id']}: {item['title'][:60]}")
        if len(not_found) > 10:
            print(f"    ... and {len(not_found) - 10} more")

    print()


if __name__ == "__main__":
    main()
