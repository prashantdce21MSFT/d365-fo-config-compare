"""Batch form comparison v2 — ALL 606 MY/SG CDDs, 4 companies each.

Reads the mapping Excel, resolves menu items from Screenshot Path / mi= / Nav Path,
compares forms across UAT (ENV1) and Config (ENV4) for MY30, MY60, SG60, DAT.
Generates a batch Excel report every 50 items.
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
    open_form, close_form, read_form_records, find_menu_item,
)
from mysg_reporter import generate_report as generate_comparison_report

# Import verified mi= values from screenshot analysis
from generate_cdd_excel import SCREENSHOT_FINDINGS

# ── Config ────────────────────────────────────────────────────────────
CONFIG_FILE = r"C:\D365DataValidator\config.json"
EXCEL_FILE = r"C:\D365 Configuration Drift Analysis\output\MYSG_CDD_FormPaths_20260428_230658.xlsx"
OUTPUT_DIR = r"C:\D365 Configuration Drift Analysis\output"
CHECKPOINT_FILE = os.path.join(OUTPUT_DIR, "batch_v2_checkpoint.json")
MI_CACHE_FILE = os.path.join(OUTPUT_DIR, "mi_cache_v2.json")

SOURCE_ENV_KEY = "ENV1"
TARGET_ENV_KEY = "ENV4"
SOURCE_DISPLAY = "Env1 UAT Asia"
TARGET_DISPLAY = "Env4 Config"

COMPANIES = ["MY30", "MY60", "SG60", "DAT"]
BATCH_SIZE = 50

# ADO IDs that are not D365 forms (emails, Excel, docs) — skip entirely
NA_ADO_IDS = set()
for ado_id, info in SCREENSHOT_FINDINGS.items():
    if info.get("mi") == "N/A" or info.get("form_type") == "N/A":
        NA_ADO_IDS.add(str(ado_id))


# ── Load Excel ────────────────────────────────────────────────────────
def load_excel_items():
    """Load all rows from the CDD mapping Excel."""
    try:
        import openpyxl
    except ImportError:
        import subprocess
        subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl"], capture_output=True)
        import openpyxl

    wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    ws = wb["CDD Form Paths"]

    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        items.append({
            "ado_id": str(int(row[0])) if isinstance(row[0], (int, float)) else str(row[0]),
            "title": row[1] or "",
            "area_path": row[2] or "",
            "state": row[3] or "",
            "company_title": row[4] or "",
            "nav_path": row[5] or "",
            "path_source": row[6] or "",
            "has_screenshots": str(row[7]).lower() == "yes" if row[7] else False,
            "screenshot_count": int(row[8]) if row[8] else 0,
            "screenshot_files": row[9] or "",
            "screenshot_path": row[10] or "",
            "mi_value": row[11] or "",
            "form_type": row[12] or "",
            "company_ss": row[13] or "",
            "notes": row[14] or "",
        })
    wb.close()
    return items


# ── MI cache ──────────────────────────────────────────────────────────
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


# ── Checkpoint ────────────────────────────────────────────────────────
def load_checkpoint():
    if os.path.isfile(CHECKPOINT_FILE):
        try:
            with open(CHECKPOINT_FILE, "r") as f:
                data = json.load(f)
            return data
        except Exception:
            pass
    return {"results": {}, "done_items": []}


def save_checkpoint(ckpt):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(CHECKPOINT_FILE, "w") as f:
        json.dump(ckpt, f, default=str)


# ── Menu item resolution ─────────────────────────────────────────────
def resolve_menu_item(client, item, mi_cache):
    """Resolve a menu item for an ADO item. Returns (menu_name, menu_type, paren_hint) or (None, None, None)."""
    ado_id = item["ado_id"]

    # Check cache first
    if ado_id in mi_cache:
        cached = mi_cache[ado_id]
        return cached.get("menu_name"), cached.get("menu_type", "Display"), cached.get("paren_hint")

    # Priority 1: SCREENSHOT_FINDINGS mi= (verified from screenshots)
    ss_finding = SCREENSHOT_FINDINGS.get(int(ado_id), {})
    if ss_finding.get("mi") and ss_finding["mi"] not in ("", "N/A"):
        mi = ss_finding["mi"]
        mi_cache[ado_id] = {"menu_name": mi, "menu_type": "Display", "paren_hint": None, "source": "screenshot_finding"}
        return mi, "Display", None

    # Priority 2: mi= Value column from Excel
    if item["mi_value"] and item["mi_value"] not in ("N/A", ""):
        mi = item["mi_value"]
        mi_cache[ado_id] = {"menu_name": mi, "menu_type": "Display", "paren_hint": None, "source": "excel_mi"}
        return mi, "Display", None

    # Priority 3: Screenshot Path (breadcrumb) — extract last segment
    path = item["screenshot_path"] or ""
    if not path or path.startswith("N/A"):
        # Priority 4: Navigation Path from title
        path = item["nav_path"] or ""

    if not path:
        mi_cache[ado_id] = {"menu_name": None, "reason": "no_path"}
        return None, None, None

    # Extract last segment from path like "Module > Setup > Form Name"
    last_segment, paren_hint = _extract_last_segment(path)
    if not last_segment:
        mi_cache[ado_id] = {"menu_name": None, "reason": "empty_last_segment"}
        return None, None, None

    # Search for menu item
    company = item.get("company_ss") or item.get("company_title") or "MY30"
    menu_name, menu_type = _search_menu_item(client, last_segment, path, paren_hint, company)

    mi_cache[ado_id] = {
        "menu_name": menu_name,
        "menu_type": menu_type or "Display",
        "paren_hint": paren_hint,
        "source": "search",
        "search_term": last_segment,
    }
    return menu_name, menu_type, paren_hint


def _extract_last_segment(path):
    """Extract the last meaningful segment from a D365 navigation path."""
    # Clean path
    clean = path.strip()
    # Remove bracket tags
    clean = re.sub(r'\[.*?\]', '', clean).strip()

    if '>' not in clean:
        # Single segment — might be "Form Name (tab hint)"
        paren_match = re.search(r'\(([^)]+)\)', clean)
        paren_hint = paren_match.group(1).strip() if paren_match else None
        base = re.sub(r'\s*\([^)]*\)', '', clean).strip()
        return base, paren_hint

    parts = [p.strip() for p in clean.split('>')]
    parts = [p for p in parts if p]
    last = parts[-1] if parts else ""

    # Extract parenthetical hint
    paren_match = re.search(r'\(([^)]+)\)', last)
    paren_hint = paren_match.group(1).strip() if paren_match else None
    base = re.sub(r'\s*\([^)]*\)', '', last).strip()

    # Strip trailing " - description"
    if ' - ' in base:
        base = base.split(' - ')[0].strip()

    # If base is empty, try parent
    if not base and len(parts) >= 2:
        base = parts[-2].strip()

    return base, paren_hint


# Generic terms that need parent context for disambiguation
_GENERIC_TERMS = {
    "parameters", "setup", "posting", "configuration", "settings",
    "number sequences", "categories", "types", "codes", "groups",
    "methods", "profiles", "journal names", "dimensions",
}

# Module hints for scoring search results
_MODULE_HINTS = {
    "cost management": ["cost", "prod"],
    "general ledger": ["ledger", "gl"],
    "accounts receivable": ["cust", "ar"],
    "accounts payable": ["vend", "ap"],
    "inventory management": ["invent"],
    "warehouse management": ["whs"],
    "procurement": ["purch"],
    "sales and marketing": ["sales", "smmparameters"],
    "transportation management": ["tms", "transport"],
    "tax": ["tax", "salestax"],
    "credit and collections": ["credit", "collection"],
    "cash and bank management": ["bank", "cash"],
    "production control": ["prod", "production"],
    "master planning": ["req", "planning"],
    "product information management": ["product", "ecores"],
    "rebate management": ["rebate"],
    "data processing framework": ["cba", "paf"],
}


def _search_menu_item(client, last_segment, full_path, paren_hint, company):
    """Search for a menu item with multi-strategy approach. Returns (name, type) or (None, None)."""
    searches = []

    # Strategy 1: paren hint
    if paren_hint and len(paren_hint) >= 3:
        searches.append(paren_hint)

    # Strategy 2: last segment
    if last_segment and len(last_segment) >= 3:
        searches.append(last_segment)

    # Strategy 3: for generic terms, combine with parent
    parts = [p.strip() for p in full_path.split('>')]
    parts = [p for p in parts if p]
    if last_segment.lower().strip() in _GENERIC_TERMS and len(parts) >= 2:
        parent = re.sub(r'\s*\[.*?\]', '', parts[-2]).strip()
        if parent:
            searches.append(f"{parent} {last_segment}")
            if len(parent) >= 3:
                searches.append(parent)

    # Strategy 4: significant words from last segment
    for word in last_segment.split():
        if len(word) >= 5 and word.lower() not in ("setup", "management", "configuration"):
            searches.append(word)

    # Deduplicate
    seen = set()
    unique = []
    for s in searches:
        k = s.lower().strip()
        if k not in seen:
            seen.add(k)
            unique.append(s)

    # Collect and score candidates
    all_candidates = []
    for search in unique:
        try:
            items = find_menu_item(client, search, company)
            for item in items:
                key = (item["name"], item["type"])
                if key not in {(c["name"], c["type"]) for c in all_candidates}:
                    all_candidates.append(item)
        except Exception:
            continue

    if not all_candidates:
        return None, None

    # Score candidates
    path_lower = full_path.lower()
    module_kws = []
    for part in parts[:-1]:
        part_l = part.lower().strip()
        for mod_key, hints in _MODULE_HINTS.items():
            if mod_key in part_l:
                module_kws.extend(hints)
                break

    scored = []
    for item in all_candidates:
        name_l = item["name"].lower()
        text_l = item["text"].lower()
        score = 0

        if text_l == last_segment.lower():
            score += 100
        elif last_segment.lower() in text_l:
            score += 50

        if paren_hint:
            if text_l == paren_hint.lower():
                score += 120
            elif paren_hint.lower() in text_l:
                score += 60

        for kw in module_kws:
            if kw in name_l:
                score += 30
            if kw in text_l:
                score += 15

        if len(item["name"]) <= 3:
            score -= 20

        scored.append((score, item))

    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best = scored[0]

    if best_score < 5:
        return None, None

    return best["name"], best["type"]


# ── Form comparison for one item + one company ───────────────────────
def compare_form_one_company(menu_name, menu_type, paren_hint, company,
                              src_client, tgt_client):
    """Compare a form for one company across two envs. Returns (result, status)."""
    # Source
    try:
        form_result = open_form(src_client, menu_name, menu_type or "Display", company)
        src_records, columns = read_form_records(src_client, form_result, include_details=False, tab_hint=paren_hint)
        close_form(src_client)
    except Exception as e:
        try:
            close_form(src_client)
        except Exception:
            pass
        return None, f"SRC_ERROR: {e}"

    # Target
    try:
        form_result = open_form(tgt_client, menu_name, menu_type or "Display", company)
        tgt_records, _ = read_form_records(tgt_client, form_result, include_details=False, tab_hint=paren_hint)
        close_form(tgt_client)
    except Exception as e:
        try:
            close_form(tgt_client)
        except Exception:
            pass
        return None, f"TGT_ERROR: {e}"

    if not src_records and not tgt_records:
        return None, "NO_DATA"

    # Key fields from first 2 columns (for grid forms)
    if columns:
        key_fields = columns[:2] if len(columns) > 2 else columns[:1]
    else:
        key_fields = []

    result = compare_entity_data(src_records, tgt_records, key_fields)
    return result, "OK"


# ── Batch report generation ──────────────────────────────────────────
def generate_batch_report(batch_results, batch_num, output_dir):
    """Generate Excel report for a batch of results."""
    if not batch_results:
        return None

    report_path = generate_comparison_report(
        batch_results, SOURCE_DISPLAY, TARGET_DISPLAY, output_dir,
        report_tag=f"Batch{batch_num:02d}"
    )
    return report_path


# ── Main ──────────────────────────────────────────────────────────────
def main():
    print("=" * 70)
    print("  MY/SG BATCH FORM COMPARISON v2 — 4 COMPANIES PER ITEM")
    print(f"  {SOURCE_DISPLAY} vs {TARGET_DISPLAY}")
    print("=" * 70)

    # 1. Load Excel items
    print("\n  Loading Excel...")
    items = load_excel_items()
    print(f"  Loaded {len(items)} items from Excel")

    # 2. Load config and connect
    with open(CONFIG_FILE) as f:
        config = json.load(f)

    print("\n  Connecting to environments...")
    src_client = D365McpClient(config["environments"][SOURCE_ENV_KEY])
    tgt_client = D365McpClient(config["environments"][TARGET_ENV_KEY])
    src_client.connect()
    print(f"  [OK] {SOURCE_DISPLAY}")
    tgt_client.connect()
    print(f"  [OK] {TARGET_DISPLAY}")

    # 3. Load checkpoint + MI cache
    ckpt = load_checkpoint()
    results = ckpt.get("results", {})
    done_items = set(ckpt.get("done_items", []))
    mi_cache = load_mi_cache()

    total = len(items)
    errors = []
    no_menu_list = []
    start_time = time.time()
    processed = 0
    batch_results = {}  # accumulate for current batch
    current_batch_start = 0

    # Figure out which batch we're in based on done count
    items_to_process = [it for it in items if it["ado_id"] not in done_items]
    done_count_at_start = len(done_items)
    # Current batch number = how many full batches already done + 1
    current_batch_num = (done_count_at_start // BATCH_SIZE) + 1
    items_in_current_batch = done_count_at_start % BATCH_SIZE

    print(f"\n  {'='*70}")
    print(f"  Total: {total} | Already done: {len(done_items)} | Remaining: {len(items_to_process)}")
    print(f"  Starting from batch {current_batch_num} (item {done_count_at_start + 1})")
    print(f"  {'='*70}\n")

    for i, item in enumerate(items):
        ado_id = item["ado_id"]

        if ado_id in done_items:
            continue

        processed += 1
        global_idx = done_count_at_start + processed

        # ETA
        elapsed = time.time() - start_time
        if processed > 1:
            eta = (elapsed / (processed - 1)) * (len(items_to_process) - processed)
            eta_str = f"{eta/60:.1f}min"
        else:
            eta_str = "..."

        print(f"\n  [{global_idx}/{total}] ADO {ado_id}: {item['title'][:80]}")

        # Skip N/A items (emails, Excel, docs)
        if ado_id in NA_ADO_IDS:
            print(f"    >> SKIP (not a D365 form)")
            done_items.add(ado_id)
            continue

        # Resolve menu item
        menu_name, menu_type, paren_hint = resolve_menu_item(src_client, item, mi_cache)

        if not menu_name:
            print(f"    >> SKIP (no menu item found)")
            no_menu_list.append((ado_id, item["title"]))
            done_items.add(ado_id)
            continue

        print(f"    Menu: {menu_name}  Type: {menu_type or 'Display'}  ETA: {eta_str}")

        # Determine company order: preferred company first, then remaining
        preferred = item.get("company_ss") or item.get("company_title") or ""
        company_order = list(COMPANIES)
        if preferred and preferred in company_order:
            company_order.remove(preferred)
            company_order.insert(0, preferred)

        # Try each company
        item_has_data = False
        for company in company_order:
            result_key = f"Configuration Deliverable {ado_id} [{company}]: {item['title']}"

            try:
                result, status = compare_form_one_company(
                    menu_name, menu_type, paren_hint, company,
                    src_client, tgt_client
                )

                if result:
                    result["item_id"] = ado_id
                    result["menu_item"] = menu_name
                    result["company"] = company
                    result["category"] = company
                    results[result_key] = result
                    batch_results[result_key] = result

                    s = result.get("summary", {})
                    sc = s.get("source_count", 0)
                    tc = s.get("target_count", 0)
                    diffs = s.get("differences", 0) + s.get("only_in_source", 0) + s.get("only_in_target", 0)
                    tag = "DIFF" if diffs else "MATCH"
                    print(f"    [{company}] {tag} (src:{sc} tgt:{tc} diffs:{diffs})")
                    item_has_data = True
                else:
                    if "ERROR" in status:
                        print(f"    [{company}] {status[:60]}")
                        errors.append((ado_id, item["title"], company, status))
                    else:
                        # NO_DATA — don't print for every company, just skip quietly
                        pass

            except Exception as e:
                print(f"    [{company}] ERROR: {e}")
                errors.append((ado_id, item["title"], company, str(e)))
                traceback.print_exc()

            time.sleep(0.2)

        if not item_has_data:
            print(f"    >> No data in any company")

        done_items.add(ado_id)

        # Checkpoint every 5 items
        if processed % 5 == 0:
            ckpt["results"] = results
            ckpt["done_items"] = list(done_items)
            save_checkpoint(ckpt)
            save_mi_cache(mi_cache)
            done_pct = len(done_items) / total * 100
            print(f"\n  --- Checkpoint ({len(done_items)}/{total}, {done_pct:.0f}%) ---\n")

        # Batch report every BATCH_SIZE items
        items_in_current_batch += 1
        if items_in_current_batch >= BATCH_SIZE:
            # Generate batch report
            if batch_results:
                print(f"\n  >>> Generating Batch {current_batch_num} report ({len(batch_results)} comparisons)...")
                try:
                    rpath = generate_batch_report(batch_results, current_batch_num, OUTPUT_DIR)
                    if rpath:
                        print(f"  >>> Saved: {rpath}")
                except Exception as e:
                    print(f"  >>> Report generation error: {e}")

            batch_results = {}
            current_batch_num += 1
            items_in_current_batch = 0

    # Final save
    ckpt["results"] = results
    ckpt["done_items"] = list(done_items)
    save_checkpoint(ckpt)
    save_mi_cache(mi_cache)

    # Generate final partial batch report
    if batch_results:
        print(f"\n  >>> Generating Batch {current_batch_num} report ({len(batch_results)} comparisons)...")
        try:
            rpath = generate_batch_report(batch_results, current_batch_num, OUTPUT_DIR)
            if rpath:
                print(f"  >>> Saved: {rpath}")
        except Exception as e:
            print(f"  >>> Report generation error: {e}")

    # Stats
    elapsed = time.time() - start_time
    total_comparisons = len(results)
    with_diffs = sum(
        1 for r in results.values()
        if isinstance(r, dict)
        and (r.get("summary", {}).get("differences", 0) > 0
             or r.get("summary", {}).get("only_in_source", 0) > 0
             or r.get("summary", {}).get("only_in_target", 0) > 0)
    )
    matches = sum(
        1 for r in results.values()
        if isinstance(r, dict)
        and r.get("summary", {}).get("differences", 0) == 0
        and r.get("summary", {}).get("only_in_source", 0) == 0
        and r.get("summary", {}).get("only_in_target", 0) == 0
        and r.get("summary", {}).get("source_count", 0) > 0
    )

    print(f"\n  {'='*70}")
    print(f"  BATCH FORM COMPARISON v2 COMPLETE")
    print(f"  {'='*70}")
    print(f"  Total ADO items: {total}")
    print(f"  Items processed: {len(done_items)}")
    print(f"  Total comparisons (item x company): {total_comparisons}")
    print(f"  Matches: {matches}")
    print(f"  With differences: {with_diffs}")
    print(f"  No menu item: {len(no_menu_list)}")
    print(f"  Errors: {len(errors)}")
    print(f"  Time: {elapsed:.0f}s ({elapsed/60:.1f} min)")

    if no_menu_list:
        print(f"\n  Items without menu item ({len(no_menu_list)}):")
        for aid, title in no_menu_list[:20]:
            print(f"    {aid}: {title[:80]}")
        if len(no_menu_list) > 20:
            print(f"    ... and {len(no_menu_list) - 20} more")

    if errors:
        print(f"\n  Errors ({len(errors)}):")
        for aid, title, comp, err in errors[:20]:
            print(f"    {aid} [{comp}]: {title[:50]} -- {err[:50]}")

        error_log = os.path.join(OUTPUT_DIR, "batch_v2_errors.txt")
        with open(error_log, "w", encoding="utf-8") as f:
            f.write(f"Batch v2 Errors\n{'='*60}\n\n")
            for aid, title, comp, err in errors:
                f.write(f"  {aid} [{comp}]: {title}\n    {err}\n\n")

    # Summary file
    summary_file = os.path.join(OUTPUT_DIR, "batch_v2_summary.txt")
    with open(summary_file, "w", encoding="utf-8") as f:
        f.write(f"MY/SG Batch Form Comparison v2 Summary\n{'='*60}\n\n")
        f.write(f"Total ADO items: {total}\n")
        f.write(f"Items processed: {len(done_items)}\n")
        f.write(f"Total comparisons: {total_comparisons}\n")
        f.write(f"Matches: {matches}\n")
        f.write(f"Diffs: {with_diffs}\n")
        f.write(f"No menu item: {len(no_menu_list)}\n")
        f.write(f"Errors: {len(errors)}\n")
        f.write(f"Time: {elapsed:.0f}s\n")


if __name__ == "__main__":
    main()
