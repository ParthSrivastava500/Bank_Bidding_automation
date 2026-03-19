"""Quick test - opens Chrome with profile and navigates to Google"""
from playwright.sync_api import sync_playwright
import time

CHROME_USER_DATA = r"C:\Users\sriva\AppData\Local\Google\Chrome\User Data"
CHROME_PROFILE   = "Profile 11"
CHROME_EXE       = "C:/Program Files/Google/Chrome/Application/chrome.exe"

with sync_playwright() as p:
    print("Launching browser...")
    context = p.chromium.launch_persistent_context(
        user_data_dir=CHROME_USER_DATA,
        executable_path=CHROME_EXE,
        headless=False,
        args=[
            f"--profile-directory={CHROME_PROFILE}",
            "--start-maximized",
            "--no-first-run",
            "--no-default-browser-check",
        ],
        ignore_default_args=[
            "--disable-extensions",
            "--disable-component-extensions-with-background-pages",
            "--disable-default-apps",
            "--disable-background-networking",
        ],
        no_viewport=True,
    )

    print(f"Browser launched! Pages: {len(context.pages)}")
    time.sleep(3)

    page = context.pages[0] if context.pages else context.new_page()
    print("Navigating to Google...")
    page.goto("https://www.google.co.in", timeout=30000)
    print(f"Success! URL: {page.url}")

    time.sleep(3)
    context.close()
    print("Done!")