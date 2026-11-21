import os
import requests
import re
from playwright.sync_api import sync_playwright

DOWNLOAD_DIR = r"c:\Users\Admin\Desktop\tester\official_ms_assets"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

urls = [
    # 1. Outlook Scheduling Assistant
    ("outlook_scheduling", "https://support.microsoft.com/en-us/office/use-the-scheduling-assistant-and-room-finder-for-meetings-in-outlook-2c8f85cb-42a1-4328-b0a6-c875d720c74b"),
    # 2. Outlook Create and Send Email
    ("outlook_compose", "https://support.microsoft.com/en-us/office/create-and-send-email-in-outlook-64fd5d4e-e184-48e8-90b9-5a1d74653a06"),
    # 3. Microsoft 365 Profile Cards & Org Chart
    ("profile_card_ms", "https://support.microsoft.com/en-us/office/profile-cards-in-microsoft-365-e80f931f-5fc4-4a5f-8707-f18d77a61677"),
    # 4. Microsoft Teams Chat
    ("teams_chat_ms", "https://support.microsoft.com/en-us/office/first-things-to-know-about-chats-in-microsoft-teams-88ff6760-330e-4293-8cfc-ef539542da44"),
    # 5. OneDrive Share Files
    ("onedrive_share_ms", "https://support.microsoft.com/en-us/office/share-onedrive-files-and-folders-9fcc2f7d-de0c-4cec-93b0-a82024800c07")
]

print("Launching Playwright to capture official Microsoft Support UI...")
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(viewport={"width": 1400, "height": 900})

    for name, url in urls:
        print(f"\nVisiting: {name} -> {url}")
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=25000)
            page.wait_for_timeout(3000)
            
            # Find all image elements on page
            img_elements = page.query_selector_all("main img, article img, .supArticleContent img")
            print(f" - Found {len(img_elements)} content images on {name}")
            
            count = 0
            for i, img in enumerate(img_elements):
                src = img.get_attribute("src")
                if src and ("support.content.office.net" in src or "support.microsoft.com" in src or "media" in src or ".png" in src or ".jpg" in src):
                    count += 1
                    try:
                        box = img.bounding_box()
                        if box and box["width"] > 100 and box["height"] > 60:
                            save_path = os.path.join(DOWNLOAD_DIR, f"{name}_{count}.png")
                            img.screenshot(path=save_path)
                            print(f"   [Saved] {save_path} ({int(box['width'])}x{int(box['height'])})")
                    except Exception as e:
                        pass
                if count >= 3:
                    break

            # Also take full section screenshot if available
            try:
                main_sec = page.query_selector("main") or page.query_selector("article")
                if main_sec:
                    sec_path = os.path.join(DOWNLOAD_DIR, f"{name}_page.png")
                    page.screenshot(path=sec_path)
                    print(f"   [Page Screenshot] {sec_path}")
            except Exception:
                pass

        except Exception as e:
            print(f" - Error visiting {name}: {e}")

    browser.close()

print("\nFinished crawling official Microsoft Support assets!")
print("Saved files in:", DOWNLOAD_DIR)
