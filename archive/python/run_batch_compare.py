"""
Batch runner: compare multiple forms between UAT (ENV1) and Config (ENV4).
"""
import sys
import os
import traceback
from datetime import datetime

sys.path.insert(0, r"C:\D365DataValidator")
sys.path.insert(0, r"C:\D365 Configuration Drift Analysis")

from form_control_extractor import FormControlExtractor
from compare_form_controls import to_excel, collect_grid_data

LEGAL_ENTITIES = ["MY30", "MY60", "SG60"]

FORMS = [
    {
        "mi_name": "InventLocations",
        "mi_type": "Display",
        "ado_id": 47860,
        "area_path": r"1875-SmartCore-ASIA\Finance and operation\EMP",
        "nav": "Inventory management > Setup > Inventory breakdown > Warehouses",
    },
    {
        "mi_name": "CHBEmptiesFGGroups",
        "mi_type": "Display",
        "ado_id": 47862,
        "area_path": r"1875-SmartCore-ASIA\Finance and operation\EMP",
        "nav": "Inventory management > Setup > Acme HuB > Empties management > Empties FG groups",
    },
    {
        "mi_name": "CHBEmptiesGroups",
        "mi_type": "Display",
        "ado_id": 47863,
        "area_path": r"1875-SmartCore-ASIA\Finance and operation\EMP",
        "nav": "Inventory management > Setup > Acme HuB > Empties management > Empties groups",
    },
    {
        "mi_name": "CBAEmptiesAlertConfiguration",
        "mi_type": "Display",
        "ado_id": 56328,
        "area_path": r"1875-SmartCore-ASIA\Finance and operation\EMP",
        "nav": "Inventory management > Setup > Acme Asia > Empties management > Empties alert > Empties alert configuration",
    },
    {
        "mi_name": "CustTableListPage",
        "mi_type": "Display",
        "ado_id": 47859,
        "area_path": r"1875-SmartCore-ASIA\Finance and operation\EMP",
        "nav": "Accounts receivable > Customers > All customers",
    },
]


def run_one(form_info):
    mi_name = form_info["mi_name"]
    mi_type = form_info["mi_type"]
    ado_id = form_info["ado_id"]
    area_path = form_info["area_path"]

    print("=" * 70)
    print(f"  Form:  {mi_name} (ADO {ado_id})")
    print(f"  Nav:   {form_info['nav']}")
    print(f"  LEs:   {', '.join(LEGAL_ENTITIES)}")
    print("=" * 70)

    # Extract from UAT (ENV1)
    print(">>> Extracting from UAT (ENV1)...")
    uat_ext = FormControlExtractor(env_key="ENV1")
    uat_result = uat_ext.extract(mi_name, LEGAL_ENTITIES, mi_type)

    # Extract from Config (ENV4)
    print("\n>>> Extracting from Config (ENV4)...")
    cfg_ext = FormControlExtractor(env_key="ENV4")
    cfg_result = cfg_ext.extract(mi_name, LEGAL_ENTITIES, mi_type)

    # Build env results list and export — just dump all values
    print("\n>>> Writing Excel...")
    env_results = [("UAT", uat_result), ("Config", cfg_result)]
    filepath = to_excel(env_results, LEGAL_ENTITIES, uat_result, cfg_result,
                        ado_id=ado_id, area_path=area_path)
    print(f"  Done! -> {filepath}\n")
    return filepath


def main():
    print("\n" + "#" * 70)
    print(f"  BATCH COMPARISON: {len(FORMS)} forms")
    print(f"  Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("#" * 70 + "\n")

    results = []
    for i, form_info in enumerate(FORMS, 1):
        print(f"\n>>> [{i}/{len(FORMS)}] {form_info['mi_name']}")
        try:
            fp = run_one(form_info)
            results.append((form_info["mi_name"], "OK", fp))
        except Exception as e:
            print(f"  ERROR: {e}")
            traceback.print_exc()
            results.append((form_info["mi_name"], f"FAILED: {e}", ""))

    print("\n" + "#" * 70)
    print("  BATCH COMPLETE")
    print("#" * 70)
    for mi, status, fp in results:
        print(f"  {mi}: {status}")
        if fp:
            print(f"    -> {fp}")
    print(f"\n  Finished: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


if __name__ == "__main__":
    main()
