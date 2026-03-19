"""
Run this once to download the correct ChromeDriver for Chrome 145
and save it to the Bidding folder.
"""
import urllib.request
import zipfile
import os
import json

print("Fetching ChromeDriver for Chrome 145...")

# Chrome for Testing API - gets exact driver for Chrome 145
url = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"

try:
    with urllib.request.urlopen(url) as r:
        data = json.loads(r.read())

    # Find latest 145.x version
    versions = [v for v in data["versions"] if v["version"].startswith("145.")]
    if not versions:
        print("No 145.x version found, trying latest...")
        versions = data["versions"]
    
    target = versions[-1]
    print(f"Using version: {target['version']}")

    # Find win64 chromedriver download
    driver_url = None
    for dl in target["downloads"].get("chromedriver", []):
        if dl["platform"] == "win64":
            driver_url = dl["url"]
            break

    if not driver_url:
        print("ERROR: Could not find win64 chromedriver download")
        exit(1)

    print(f"Downloading from: {driver_url}")
    zip_path = "chromedriver_145.zip"
    urllib.request.urlretrieve(driver_url, zip_path)
    print("Downloaded!")

    # Extract chromedriver.exe
    with zipfile.ZipFile(zip_path, 'r') as z:
        for name in z.namelist():
            if name.endswith("chromedriver.exe"):
                # Extract to current folder
                with z.open(name) as src, open("chromedriver.exe", "wb") as dst:
                    dst.write(src.read())
                print(f"Extracted: chromedriver.exe")
                break

    os.remove(zip_path)
    print("\n✅ Done! chromedriver.exe is ready in your Bidding folder.")
    print("Now run: python main.py")

except Exception as e:
    print(f"Error: {e}")
    print("\nManual download:")
    print("1. Go to: https://googlechromelabs.github.io/chrome-for-testing/")
    print("2. Find version 145.x → chromedriver → win64")
    print("3. Download, extract chromedriver.exe to your Bidding folder")
