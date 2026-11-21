import requests
import re

u = "https://support.microsoft.com/en-us/office/use-the-scheduling-assistant-and-room-finder-for-meetings-in-outlook-7a00e4ab-f7ae-4bef-b428-fabc973a0e82"
r = requests.get(u, headers={"User-Agent": "Mozilla/5.0"})
all_urls = re.findall(r'https?://[^\s"\'<>\)]+\.(?:png|jpg|jpeg|webp)', r.text, re.IGNORECASE)
print("Total image URLs in HTML:", len(all_urls))
for x in list(set(all_urls))[:15]:
    print(" -", x)
