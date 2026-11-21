import asyncio
from playwright.async_api import async_playwright

async def inspect_booking_modal():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        await page.goto('https://uat.inxify.com/login')
        await page.fill('input#username', 'quanhm2')
        await page.fill('input#password', '@Matkhau1')
        await page.click('button[type="submit"]')
        await page.wait_for_timeout(3000)
        
        await page.goto('https://uat.inxify.com/pms/reservations')
        await page.wait_for_timeout(3000)
        
        # Kill popups
        await page.evaluate('''() => {
            document.querySelectorAll('.ant-tour, .ant-tour-mask, .ant-tour-root').forEach(e => e.remove());
        }''')
        
        # Click "+ Đặt phòng mới"
        await page.click('button:has-text("+ Đặt phòng mới"), button:has-text("Đặt phòng mới")')
        await page.wait_for_timeout(2000)
        
        fields = await page.evaluate('''() => {
            const modal = document.querySelector('.ant-modal-body, .ant-modal-content');
            if (!modal) return 'No modal';
            const labels = Array.from(modal.querySelectorAll('label, .ant-form-item-label')).map(l => l.innerText.trim());
            const inputs = Array.from(modal.querySelectorAll('input, select, textarea')).map(i => ({
                id: i.id,
                placeholder: i.placeholder,
                type: i.type,
                value: i.value
            }));
            return { text: modal.innerText, labels, inputs };
        }''')
        print("Labels:", fields['labels'])
        print("Inputs:", fields['inputs'])
        
        await browser.close()

if __name__ == '__main__':
    asyncio.run(inspect_booking_modal())
