import requests
import re
import os

OUT_DIR = r"c:\Users\Admin\Desktop\tester\official_ms_assets"
os.makedirs(OUT_DIR, exist_ok=True)

urls = [
    "https://support.microsoft.com/en-us/office/use-the-scheduling-assistant-and-room-finder-for-meetings-in-outlook-7a00e4ab-f7ae-4bef-b428-fabc973a0e82",
    "https://support.microsoft.com/en-us/office/create-and-send-email-in-outlook-64fd5d4e-e184-48e8-90b9-5a1d74653a06",
    "https://support.microsoft.com/en-us/office/profile-cards-in-microsoft-365-e80f931f-5fc4-4a5f-8707-f18d77a61677",
    "https://support.microsoft.com/en-us/office/share-onedrive-files-and-folders-9fcc2f7d-de0c-4cec-93b0-a82024800c07",
    "https://support.microsoft.com/en-us/office/first-things-to-know-about-chats-in-microsoft-teams-88ff6760-330e-4293-8cfc-ef539542da44"
]

all_img_urls = []
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

for u in urls:
    try:
        r = requests.get(u, headers=headers, timeout=10)
        found = re.findall(r'https://support\.content\.office\.net/[^\s"\'<>\)]+', r.text)
        print(f"URL: {u.split('/')[-1][:25]} -> Found {len(found)} images")
        for f in found:
            # clean up any trailing punctuation
            clean_url = f.rstrip(".,;\"'")
            if any(ext in clean_url.lower() for ext in [".png", ".jpg", ".jpeg", ".gif"]):
                all_img_urls.append(clean_url)
    except Exception as e:
        print("Error fetching:", u, e)

unique_imgs = list(set(all_img_urls))
print(f"Total unique image URLs on Microsoft CDN: {len(unique_imgs)}")

for idx, img_url in enumerate(unique_imgs):
    try:
        r = requests.get(img_url, headers=headers, timeout=10)
        if r.status_code == 200 and len(r.content) > 10000:
            fn = os.path.basename(img_url.split("?")[0])
            if not fn.endswith(".png"):
                fn = f"ms_img_{idx}.png"
            dest = os.path.join(OUT_DIR, fn)
            with open(dest, "wb") as fp:
                fp.write(r.content)
            print(f" [Downloaded] {fn} ({len(r.content)} bytes) from {img_url[:60]}")
    except Exception as e:
        pass
