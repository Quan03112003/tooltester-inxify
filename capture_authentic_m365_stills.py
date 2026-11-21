import os
from playwright.sync_api import sync_playwright

HTML_PATH = r"c:\Users\Admin\Desktop\tester\create_video_animation.html"
IMG_DIR = r"c:\Users\Admin\Desktop\tester\img_m365_guide"
os.makedirs(IMG_DIR, exist_ok=True)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(viewport={"width": 1280, "height": 720}, device_scale_factor=2)
    page.goto(f"file:///{HTML_PATH.replace(os.sep, '/')}")
    page.wait_for_timeout(1000)

    # 1. Teams & Profile Card
    page.evaluate("""
        document.querySelectorAll('.scene').forEach(s => s.classList.remove('active'));
        document.getElementById('scene-teams').classList.add('active');
        document.getElementById('profile-pop').classList.add('visible');
        document.getElementById('virtual-cursor').style.left = '460px';
        document.getElementById('virtual-cursor').style.top = '275px';
    """)
    page.wait_for_timeout(500)
    page.screenshot(path=os.path.join(IMG_DIR, "M365_03_ProfileCard_OrgChart.png"))
    print("Captured authentic M365_03_ProfileCard_OrgChart.png")

    # 2. Authentic Outlook Desktop (3-Column Layout with Compose Modal)
    page.evaluate("""
        document.querySelectorAll('.scene').forEach(s => s.classList.remove('active'));
        document.getElementById('scene-outlook').classList.add('active');
        document.getElementById('modal-compose').classList.add('visible');
        document.getElementById('modal-sched').classList.remove('visible');
        document.getElementById('virtual-cursor').style.left = '80px';
        document.getElementById('virtual-cursor').style.top = '110px';
    """)
    page.wait_for_timeout(500)
    page.screenshot(path=os.path.join(IMG_DIR, "M365_01_SoSanh_Zalo_vs_M365.png"))
    print("Captured authentic M365_01_SoSanh_Zalo_vs_M365.png (Outlook Desktop & Compose Modal)")

    # 3. Authentic Outlook Scheduling Assistant
    page.evaluate("""
        document.querySelectorAll('.scene').forEach(s => s.classList.remove('active'));
        document.getElementById('scene-outlook').classList.add('active');
        document.getElementById('modal-compose').classList.remove('visible');
        document.getElementById('modal-sched').classList.add('visible');
        document.getElementById('virtual-cursor').style.left = '680px';
        document.getElementById('virtual-cursor').style.top = '280px';
    """)
    page.wait_for_timeout(500)
    page.screenshot(path=os.path.join(IMG_DIR, "M365_04_TrungTam_Video_DaoTao.png"))
    print("Captured authentic M365_04_TrungTam_Video_DaoTao.png (Outlook Scheduling Assistant)")

    # 4. Matrix 4 Tools & Sunset Zalo
    page.evaluate("""
        document.querySelectorAll('.scene').forEach(s => s.classList.remove('active'));
        document.getElementById('scene-matrix').classList.add('active');
        document.getElementById('virtual-cursor').style.left = '-100px';
    """)
    page.wait_for_timeout(500)
    page.screenshot(path=os.path.join(IMG_DIR, "M365_02_PhanDinh_4_CongCu.png"))
    print("Captured authentic M365_02_PhanDinh_4_CongCu.png (4-Tool Boundary Matrix)")

    browser.close()

print("All authentic screenshots updated!")
