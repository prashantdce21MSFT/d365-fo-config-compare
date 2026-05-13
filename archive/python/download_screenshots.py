"""Download screenshots from ADO items that need mi= extraction."""
import json
import os
import sys
import urllib.request
import time

OUTPUT_DIR = r"C:\D365 Configuration Drift Analysis\output\screenshots"

def get_token():
    token = os.environ.get("ADO_TOKEN")
    if token:
        return token
    import subprocess
    r = subprocess.run(
        ["az", "account", "get-access-token", "--resource",
         "499b84ac-1321-427f-aa17-267ca6975798", "--query", "accessToken", "-o", "tsv"],
        capture_output=True, text=True, shell=True
    )
    return r.stdout.strip()

def download_image(url, token, save_path):
    """Download image from ADO, resize if needed."""
    req = urllib.request.Request(url)
    req.add_header("Authorization", f"Bearer {token}")
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = resp.read()
    with open(save_path, "wb") as f:
        f.write(data)
    
    # Resize if too large (>1600px) for Claude Read tool
    try:
        from PIL import Image
        img = Image.open(save_path)
        w, h = img.size
        if max(w, h) > 1600:
            ratio = 1600 / max(w, h)
            new_size = (int(w * ratio), int(h * ratio))
            img = img.resize(new_size, Image.LANCZOS)
            img.save(save_path)
    except ImportError:
        pass
    return save_path

def main():
    token = get_token()
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    with open(r"C:\D365 Configuration Drift Analysis\output\all_ado_items.json") as f:
        items = json.load(f)
    
    # Download screenshots for items WITHOUT nav path (48 items) 
    # AND for items WITH path AND screenshots (to get mi= values)
    targets = [i for i in items if i["has_screenshots"]]
    
    print(f"Total items with screenshots: {len(targets)}")
    downloaded = 0
    errors = 0
    
    for item in targets:
        ado_id = item["ado_id"]
        urls = item["screenshot_urls"]
        
        for idx, url in enumerate(urls):
            fname = f"{ado_id}_img{idx+1}.png"
            fpath = os.path.join(OUTPUT_DIR, fname)
            
            if os.path.exists(fpath):
                downloaded += 1
                continue
            
            try:
                # Fix relative URLs
                if url.startswith("/"):
                    url = f"https://dev.azure.com/Acmegroup{url}"
                elif not url.startswith("http"):
                    url = f"https://dev.azure.com/Acmegroup/1875-SmartCore-ASIA{url}"
                
                download_image(url, token, fpath)
                downloaded += 1
                if downloaded % 20 == 0:
                    sys.stdout.write(f"\r  Downloaded: {downloaded}")
                    sys.stdout.flush()
                time.sleep(0.2)
            except Exception as e:
                errors += 1
                if errors <= 5:
                    print(f"\n  Error {ado_id}_img{idx+1}: {e}")
    
    print(f"\n  Done: {downloaded} downloaded, {errors} errors")
    
    # Count total files
    files = os.listdir(OUTPUT_DIR)
    print(f"  Total screenshot files: {len(files)}")

if __name__ == "__main__":
    main()
