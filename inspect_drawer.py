import asyncio
from playwright.async_api import async_playwright

async def inspect_drawer():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(viewport={'width': 1920, 'height': 1080})
        
        await page.goto('https://uat.inxify.com/login')
        await page.fill('input#username', 'quanhm2')
        await page.fill('input#password', '@Matkhau1')
        await page.click('button[type="submit"]')
        await page.wait_for_timeout(3000)
        
        await page.goto('https://uat.inxify.com/crm/guests')
        await page.wait_for_timeout(3000)
        
        # Close any modal
        await page.evaluate('''() => {
            document.querySelectorAll('.ant-modal-root, .ant-modal-mask, .ant-modal-wrap, .ant-tour').forEach(e => e.remove());
        }''')
        await page.wait_for_timeout(1000)
        
        # Click on the first "Xem" button
        xem_buttons = page.locator('button:has-text("Xem")')
        count = await xem_buttons.count()
        print(f"Found {count} 'Xem' buttons")
        if count > 0:
            await xem_buttons.first.click()
            await page.wait_for_timeout(2000)
            
            # Check what opened
            drawer_content = await page.evaluate('''() => {
                const drawer = document.querySelector('.ant-drawer-content, .ant-drawer');
                if (!drawer) return 'No drawer found';
                return {
                    text: drawer.innerText,
                    buttons: Array.from(drawer.querySelectorAll('button')).map(b => b.innerText.trim()).filter(Boolean),
                    inputs: Array.from(drawer.querySelectorAll('input, select, textarea')).map(i => ({
                        type: i.type,
                        value: i.value,
                        placeholder: i.placeholder,
                        disabled: i.disabled,
                        readOnly: i.readOnly
                    }))
                };
            }''')
            print("Drawer info:", drawer_content)
        
        # Check if there are other buttons on CRM header or subheader
        header_text = await page.evaluate('''() => {
            const h = document.querySelector('.ant-page-header, main');
            return h ? h.innerText.substring(0, 500) : '';
        }''')
        print("Header area:", header_text)

        await browser.close()

if __name__ == '__main__':
    asyncio.run(inspect_drawer())
