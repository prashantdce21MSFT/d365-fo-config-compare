"""Fetch all ADO items from MY/SG query and extract form paths."""
import json
import re
import subprocess
import time
import sys
import os

QUERY_ID = "7a10e4a5-21aa-4c0d-b9e6-82e88294db0f"
ORG = "https://dev.azure.com/Acmegroup"
PROJECT = "1875-SmartCore-ASIA"
OUTPUT_DIR = r"C:\D365 Configuration Drift Analysis\output"

def get_token():
    token = os.environ.get("ADO_TOKEN")
    if token:
        return token
    r = subprocess.run(
        ["az", "account", "get-access-token", "--resource",
         "499b84ac-1321-427f-aa17-267ca6975798", "--query", "accessToken", "-o", "tsv"],
        capture_output=True, text=True, shell=True
    )
    return r.stdout.strip()

def ado_get(url, token):
    import urllib.request
    req = urllib.request.Request(url)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode())

def ado_post(url, token, body):
    import urllib.request
    data = json.dumps(body).encode()
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode())

def main():
    token = get_token()
    print(f"Token obtained: {len(token)} chars")

    # 1. Run the query to get work item IDs
    query_url = f"{ORG}/{PROJECT}/_apis/wit/wiql/{QUERY_ID}?api-version=7.0"
    print(f"Running query...")
    result = ado_get(query_url, token)
    
    wi_ids = [item["id"] for item in result.get("workItems", [])]
    print(f"Query returned {len(wi_ids)} work items")

    # 2. Fetch work items in batches of 200 (API limit)
    all_items = []
    fields = ["System.Id", "System.Title", "System.Description", "System.AreaPath", 
              "System.State", "System.IterationPath"]
    
    for batch_start in range(0, len(wi_ids), 200):
        batch_ids = wi_ids[batch_start:batch_start+200]
        print(f"  Fetching batch {batch_start//200 + 1}: IDs {batch_start+1}-{batch_start+len(batch_ids)}")
        
        batch_url = f"{ORG}/{PROJECT}/_apis/wit/workitemsbatch?api-version=7.0"
        body = {"ids": batch_ids, "fields": fields}
        
        batch_result = ado_post(batch_url, token, body)
        items = batch_result.get("value", [])
        all_items.extend(items)
        time.sleep(0.5)
    
    print(f"\nTotal items fetched: {len(all_items)}")

    # 3. Process each item
    results = []
    for item in all_items:
        fields_data = item.get("fields", {})
        ado_id = fields_data.get("System.Id", item.get("id"))
        title = fields_data.get("System.Title", "")
        desc = fields_data.get("System.Description", "") or ""
        area = fields_data.get("System.AreaPath", "")
        state = fields_data.get("System.State", "")
        
        # Extract navigation path from title (pattern: "Module > Setup > Form")
        nav_path = ""
        if ">" in title:
            # Title itself contains the path
            nav_path = title.strip()
        
        # Extract company from title [MY], [SG], etc.
        company = ""
        for m in re.finditer(r'\[([A-Za-z0-9]{2,6})\]', title):
            tag = m.group(1)
            if tag.lower() not in ("export", "import", "empties", "cutover", "fdd"):
                company = tag
                break
        
        # Find screenshot URLs in description HTML
        img_urls = re.findall(r'src="([^"]*/_apis/wit/attachments/[^"]*)"', desc)
        
        # Find any mi= in description text
        mi_matches = re.findall(r'mi=([A-Za-z0-9_]+)', desc)
        mi_from_desc = mi_matches[0] if mi_matches else ""
        
        # Find any D365 URLs in description
        d365_urls = re.findall(r'https?://[a-z0-9-]+\.(?:sandbox\.)?operations\.dynamics\.com[^\s"<>]*', desc)
        
        results.append({
            "ado_id": ado_id,
            "title": title,
            "area_path": area,
            "state": state,
            "nav_path": nav_path,
            "company": company,
            "has_screenshots": len(img_urls) > 0,
            "screenshot_count": len(img_urls),
            "screenshot_urls": img_urls,
            "mi_from_desc": mi_from_desc,
            "d365_urls": d365_urls,
        })
    
    # Save raw data
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_file = os.path.join(OUTPUT_DIR, "all_ado_items.json")
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved to {out_file}")
    
    # Stats
    with_path = sum(1 for r in results if r["nav_path"])
    with_screenshots = sum(1 for r in results if r["has_screenshots"])
    with_mi = sum(1 for r in results if r["mi_from_desc"])
    with_d365 = sum(1 for r in results if r["d365_urls"])
    with_company = sum(1 for r in results if r["company"])
    
    print(f"\n--- Stats ---")
    print(f"  Total items:          {len(results)}")
    print(f"  With nav path in title: {with_path}")
    print(f"  With screenshots:     {with_screenshots}")
    print(f"  With mi= in desc:    {with_mi}")
    print(f"  With D365 URLs:       {with_d365}")
    print(f"  With company tag:     {with_company}")

if __name__ == "__main__":
    main()
