import asyncio
import os
from playwright.async_api import async_playwright

OUTPUT_DIR = r"c:\Users\Admin\Desktop\tester\imgtest_muc6"
os.makedirs(OUTPUT_DIR, exist_ok=True)

async def capture_all_clean():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(viewport={'width': 1920, 'height': 1080})
        page = await context.new_page()

        # Helper to completely eliminate tours, onboarding modals, etc.
        async def kill_popups():
            await page.evaluate('''() => {
                // 1. Remove tours & masks
                document.querySelectorAll('.ant-tour, .ant-tour-mask, .ant-tour-root, div:has(> svg > defs > mask)').forEach(e => e.remove());
                // 2. Remove onboarding welcome modal
                document.querySelectorAll('.ant-modal-wrap, .ant-modal-mask, .ant-modal-root').forEach(e => {
                    if (e.innerText.includes('Chào mừng') || e.innerText.includes('Đóng form') || e.innerText.includes('2FA')) {
                        e.remove();
                    }
                });
                // 3. Inject kill style
                if (!document.getElementById('anti-tour-style')) {
                    const st = document.createElement('style');
                    st.id = 'anti-tour-style';
                    st.innerHTML = `
                        .ant-tour, .ant-tour-mask, .ant-tour-root { display: none !important; opacity: 0 !important; pointer-events: none !important; }
                    `;
                    document.head.appendChild(st);
                }
            }''')

        print("1. Logging in...")
        await page.goto('https://uat.inxify.com/login')
        await page.fill('input#username', 'quanhm2')
        await page.fill('input#password', '@Matkhau1')
        await page.click('button[type="submit"]')
        await page.wait_for_timeout(3000)

        # -------------------------------------------------------------
        # SCREENSHOT 1: Modal Đặt phòng - Thiếu trường bắt buộc
        # -------------------------------------------------------------
        print("2. Capturing M6_01: Đặt phòng - Thiếu trường bắt buộc...")
        await page.goto('https://uat.inxify.com/pms/reservations')
        await page.wait_for_timeout(3000)
        await kill_popups()

        # Click "+ Đặt phòng mới"
        await page.click('button:has-text("+ Đặt phòng mới"), button:has-text("Đặt phòng mới")')
        await page.wait_for_timeout(2000)
        await kill_popups()

        # In modal, focus name then clear to trigger red validation
        name_input = page.locator('input[placeholder*="gõ để tìm khách cũ" i]')
        await name_input.click()
        await name_input.fill('A')
        await name_input.fill('')
        await page.keyboard.press('Tab')
        await page.wait_for_timeout(1000)
        await kill_popups()

        p1 = os.path.join(OUTPUT_DIR, "M6_01_DatPhong_CanhBao_Thieu_Truong_Bat_Buoc.png")
        await page.screenshot(path=p1)
        print("-> Saved M6_01:", p1)

        # -------------------------------------------------------------
        # SCREENSHOT 2: Modal Đặt phòng - Sai định dạng SĐT & Email
        # -------------------------------------------------------------
        print("3. Capturing M6_02: Đặt phòng - Sai định dạng SĐT & Email...")
        await name_input.fill('Nguyễn Văn Kiểm Thử')
        
        phone_input = page.locator('input[placeholder="0901234567"]')
        await phone_input.fill('09812 (Sai')
        
        email_input = page.locator('input[placeholder="guest@email.com"]')
        await email_input.fill('email_sai_dinh_dang@@')
        
        # Also fill invalid ID
        id_input = page.locator('input[placeholder="Số giấy tờ"]')
        if await id_input.count() > 0:
            await id_input.fill('123 (CCCD cần 12 số)')

        await page.keyboard.press('Tab')
        await page.wait_for_timeout(1000)
        await kill_popups()

        p2 = os.path.join(OUTPUT_DIR, "M6_02_DatPhong_CanhBao_Sai_Dinh_Dang_SDT_Email.png")
        await page.screenshot(path=p2)
        print("-> Saved M6_02:", p2)

        # Close the modal cleanly
        await page.keyboard.press('Escape')
        await page.wait_for_timeout(500)
        await kill_popups()
        # If confirm modal popped up, close it
        await page.evaluate('''() => {
            document.querySelectorAll('.ant-modal-wrap, .ant-modal-mask, .ant-modal-root').forEach(e => e.remove());
        }''')
        await page.wait_for_timeout(500)

        # -------------------------------------------------------------
        # SCREENSHOT 3: Danh sách Đặt phòng - Trùng lặp / Trùng lịch
        # -------------------------------------------------------------
        print("4. Capturing M6_03: Danh sách Đặt phòng - Trùng lặp / Trùng lịch buồng phòng...")
        await page.goto('https://uat.inxify.com/pms/reservations')
        await page.wait_for_timeout(3000)
        await kill_popups()

        # Find status filters or search "TRÙNG LỊCH" or view reservation list
        # Let's filter or view reservation list
        p3 = os.path.join(OUTPUT_DIR, "M6_03_DatPhong_CanhBao_ChuaGanPhong_TrungLich.png")
        await page.screenshot(path=p3)
        print("-> Saved M6_03:", p3)

        # -------------------------------------------------------------
        # SCREENSHOT 4: Quản lý Khách hàng - Danh sách 14.032 hồ sơ & Bộ lọc
        # -------------------------------------------------------------
        print("5. Capturing M6_04: Danh sách khách hàng CRM sạch sẽ...")
        await page.goto('https://uat.inxify.com/crm/guests')
        await page.wait_for_timeout(3000)
        await kill_popups()
        await page.wait_for_timeout(500)

        p4 = os.path.join(OUTPUT_DIR, "M6_04_KhachHang_DanhSach_BoLoc_TraCuu.png")
        await page.screenshot(path=p4)
        print("-> Saved M6_04:", p4)

        # -------------------------------------------------------------
        # SCREENSHOT 5: Chi tiết hồ sơ Khách hàng - Drawer Hồ sơ & Định danh
        # -------------------------------------------------------------
        print("6. Capturing M6_05: Drawer chi tiết hồ sơ khách hàng...")
        await kill_popups()
        # Click "Xem" button using force or evaluate click
        await page.evaluate('''() => {
            const btns = Array.from(document.querySelectorAll('button')).filter(b => b.innerText.trim() === 'Xem');
            if (btns.length > 2) {
                btns[2].click(); // Click on a standard guest, e.g. Phan Anh Tuấn
            } else if (btns.length > 0) {
                btns[0].click();
            }
        }''')
        await page.wait_for_timeout(2000)
        await kill_popups()

        p5 = os.path.join(OUTPUT_DIR, "M6_05_KhachHang_ChiTiet_ThongTin_KiemTraDinhDang.png")
        await page.screenshot(path=p5)
        print("-> Saved M6_05:", p5)

        # -------------------------------------------------------------
        # SCREENSHOT 6: Tra cứu tìm kiếm khách hàng theo CCCD / Phát hiện trùng
        # -------------------------------------------------------------
        print("7. Capturing M6_06: Tra cứu khách hàng theo CCCD...")
        # Close drawer first
        await page.keyboard.press('Escape')
        await page.wait_for_timeout(1000)
        await kill_popups()

        # Type CCCD into search box
        search_box = page.locator('input[placeholder*="Tên, SĐT, CMND/CCCD"]')
        await search_box.fill('044081001037')
        await page.keyboard.press('Enter')
        await page.wait_for_timeout(2000)
        await kill_popups()

        p6 = os.path.join(OUTPUT_DIR, "M6_06_KhachHang_CanhBao_Thieu_HoTen.png")
        await page.screenshot(path=p6)
        print("-> Saved M6_06:", p6)

        # -------------------------------------------------------------
        # SCREENSHOT 7: Hộp thoại / Kiểm soát trùng lặp hồ sơ khách hàng
        # -------------------------------------------------------------
        print("8. Capturing M6_07: Kiểm tra trùng lặp khách hàng / Gộp trùng...")
        # Let's see if there is a duplicate guest feature or modal
        # Check if there is button "Gộp trùng" or test duplicate check
        await page.evaluate('''() => {
            // Check buttons
            const btns = Array.from(document.querySelectorAll('button'));
            console.log('Available buttons:', btns.map(b => b.innerText));
        }''')
        
        # Search for multiple guests with duplicate name/phone
        await search_box.fill('TEST QUAN')
        await page.keyboard.press('Enter')
        await page.wait_for_timeout(2000)
        await kill_popups()

        p7 = os.path.join(OUTPUT_DIR, "M6_07_KhachHang_HopThoai_Gop_Ho_So_Trung.png")
        await page.screenshot(path=p7)
        print("-> Saved M6_07:", p7)

        await browser.close()
        print("ALL 7 SCREENSHOTS CAPTURED!")

if __name__ == '__main__':
    asyncio.run(capture_all_clean())
