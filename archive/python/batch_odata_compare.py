"""Batch OData comparison: Env1 UAT Asia (ENV1) vs Env4 Config (ENV4).

Reads pre-computed categorization from cat_checkpoint.json, compares
BOTH + ODATA items via OData data entities. Checkpoints every 25 items.
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
from data_fetcher import fetch_entity_data, select_safe_fields, print_progress_bar
from comparator import compare_entity_data
from mysg_reporter import generate_report as generate_comparison_report

# ── Config ────────────────────────────────────────────────────────────
CONFIG_FILE = r"C:\D365DataValidator\config.json"
CAT_FILE = r"C:\D365 Configuration Drift Analysis\cat_checkpoint.json"
OUTPUT_DIR = r"C:\D365 Configuration Drift Analysis\output"
CHECKPOINT_FILE = os.path.join(OUTPUT_DIR, "odata_checkpoint.json")

SOURCE_ENV_KEY = "ENV1"
TARGET_ENV_KEY = "ENV4"
SOURCE_DISPLAY = "Env1 UAT Asia"
TARGET_DISPLAY = "Env4 Config"


# ── Skip list (entities too large or irrelevant) ─────────────────────
_SKIP_ENTITIES = {"BatchJobs"}

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
        elif key.startswith("ODATA_"):
            ado = key[6:].split(" ")[0]
            ids.add(ado)
        elif key.startswith("Configuration Deliverable "):
            ado = key.split(":")[0].split()[-1]
            ids.add(ado)
    return ids


def extract_key_fields(client, entity_name):
    """Get key fields and select fields from entity metadata."""
    raw = client.get_entity_metadata(entity_name, include_keys=True)
    selected, keys = select_safe_fields(raw, max_fields=25)
    return {"keys": keys, "select_fields": selected}


# ── Main ──────────────────────────────────────────────────────────────
def main():
    print("=" * 70)
    print("  MY/SG BATCH ODATA COMPARISON")
    print(f"  {SOURCE_DISPLAY} vs {TARGET_DISPLAY}")
    print("=" * 70)

    # 1. Load config
    with open(CONFIG_FILE) as f:
        config = json.load(f)

    # 2. Load categorization
    with open(CAT_FILE) as f:
        all_items = json.load(f)

    # Filter to BOTH + ODATA items (these have odata_entity populated)
    odata_items = [
        item for item in all_items
        if item.get("category") in ("BOTH", "ODATA")
        and item.get("odata_entity")
    ]
    # Also try UNCERTAIN items that have a decent odata_score
    uncertain_odata = [
        item for item in all_items
        if item.get("category") == "UNCERTAIN"
        and item.get("odata_score", 0) >= 30
        and item.get("odata_entity")
    ]
    odata_items.extend(uncertain_odata)

    print(f"  Total categorized items: {len(all_items)}")
    print(f"  BOTH+ODATA items: {len(odata_items) - len(uncertain_odata)}")
    print(f"  UNCERTAIN (trying OData): {len(uncertain_odata)}")
    print(f"  Total to process: {len(odata_items)}")

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

    max_records = config.get("settings", {}).get("max_records_per_entity", 50000)

    # 5. Process each OData item
    total = len(odata_items)
    errors = []
    start_time = time.time()

    print(f"\n  {'='*70}")
    print(f"  Starting comparison ({total} items, {len(already_done)} already done)")
    print(f"  {'='*70}\n")

    for i, item in enumerate(odata_items, 1):
        ado_id = str(item["id"])
        title = item["title"]
        entity_name = item["odata_entity"]

        if ado_id in already_done:
            skipped_count += 1
            continue

        if entity_name in _SKIP_ENTITIES:
            print(f"\n  [{i}/{total}] {ado_id}: SKIPPED (entity {entity_name} in skip list)")
            skipped_count += 1
            continue

        # Extract company from title for dataAreaId filter
        company = extract_company_from_title(title)
        le_filter = f"dataAreaId eq '{company.lower()}'" if company else None

        elapsed = time.time() - start_time
        processed = i - skipped_count
        if processed > 1:
            eta = (elapsed / (processed - 1)) * (total - i)
            eta_str = f"{eta/60:.1f}min"
        else:
            eta_str = "..."

        print(f"\n  [{i}/{total}] {ado_id}: {title[:80]}")
        print(f"    Entity: {entity_name}  Company: {company or 'cross-company'}  ETA: {eta_str}")

        try:
            # Get metadata (key fields + safe select fields)
            fields_info = extract_key_fields(src_client, entity_name)

            # Fetch source
            sys.stdout.write(f"    source: ")
            src_records = fetch_entity_data(
                src_client, entity_name,
                select_fields=fields_info.get("select_fields"),
                filter_expr=le_filter,
                max_records=max_records,
                progress_callback=print_progress_bar,
            )

            # Fetch target
            sys.stdout.write(f"    target: ")
            tgt_records = fetch_entity_data(
                tgt_client, entity_name,
                select_fields=fields_info.get("select_fields"),
                filter_expr=le_filter,
                max_records=max_records,
                progress_callback=print_progress_bar,
            )

            if not src_records and not tgt_records:
                print(f"    >> No data in either environment")
                continue

            # Compare
            result = compare_entity_data(src_records, tgt_records, fields_info["keys"])

            display_key = f"Configuration Deliverable {ado_id}: {title}"
            result["item_id"] = ado_id
            result["category"] = item.get("category", "OData")
            results[display_key] = result

            s = result.get("summary", {})
            sc = s.get("source_count", 0)
            tc = s.get("target_count", 0)
            diffs = s.get("differences", 0) + s.get("only_in_source", 0) + s.get("only_in_target", 0)
            status = "DIFF" if diffs else "MATCH"
            print(f"    >> {status} (src:{sc} tgt:{tc} diffs:{diffs})")

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

    # 6. Final save + report
    save_checkpoint(results)

    print(f"\n  {'='*70}")
    print(f"  ODATA COMPARISON COMPLETE")
    print(f"  {'='*70}")
    print(f"  Total OData items: {total}")
    print(f"  Results collected: {len(results)}")
    print(f"  Errors: {len(errors)}")

    if errors:
        print(f"\n  Failed items:")
        for ado_id, title_text, err in errors[:30]:
            print(f"    {ado_id}: {title_text[:80]} -- {err[:80]}")

        # Save error log
        error_log = os.path.join(OUTPUT_DIR, "odata_errors.txt")
        with open(error_log, "w", encoding="utf-8") as f:
            f.write(f"OData Comparison Errors\n{'='*60}\n\n")
            for ado_id, title_text, err in errors:
                f.write(f"  {ado_id}: {title_text}\n    {err}\n\n")
        print(f"  Error log: {error_log}")

    if results:
        print(f"\n  Generating Excel report...")
        report_path = generate_comparison_report(
            results, SOURCE_DISPLAY, TARGET_DISPLAY, OUTPUT_DIR, report_tag="OData"
        )
        print(f"  Report saved: {report_path}")
    else:
        print("\n  No results to report.")

    elapsed = time.time() - start_time
    print(f"\n  Total time: {elapsed:.0f}s ({elapsed/60:.1f} min)")


if __name__ == "__main__":
    main()
