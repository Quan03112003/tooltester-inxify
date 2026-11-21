import asyncio
from playwright.async_api import async_playwright

async def inspect():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(viewport={'width': 1920, 'height': 1080})
        page = await context.new_page()
        
        print("Navigating to login...")
        await page.goto('https://uat.inxify.com/login')
        await page.wait_for_timeout(2000)
        
        await page.fill('input#username', 'quanhm2')
        await page.fill('input#password', '@Matkhau1')
        await page.click('button[type="submit"]')
        await page.wait_for_timeout(3000)
        print("Logged in, URL:", page.url)
        
        # Navigate to CRM
        print("Navigating to /crm/guests...")
        await page.goto('https://uat.inxify.com/crm/guests')
        await page.wait_for_timeout(4000)
        print("CRM URL:", page.url)
        
        # Check modals and texts
        modals = await page.evaluate('''() => {
            return Array.from(document.querySelectorAll('.ant-modal-content, .ant-tour-inner, .ant-modal')).map(el => el.innerText.trim());
        }''')
        print("Modals found on CRM:", modals)
        
        # Check buttons
        buttons = await page.evaluate('''() => {
            return Array.from(document.querySelectorAll('button')).map(el => el.innerText.trim()).filter(Boolean);
        }''')
        print("Buttons found on CRM:", buttons)
        
        # Check inputs and search
        inputs = await page.evaluate('''() => {
            return Array.from(document.querySelectorAll('input')).map(el => ({
                placeholder: el.placeholder,
                type: el.type,
                name: el.name,
                id: el.id
            }));
        }''')
        print("Inputs found on CRM:", inputs)
        
        # Check table headers and rows
        table_headers = await page.evaluate('''() => {
            return Array.from(document.querySelectorAll('th')).map(el => el.innerText.trim()).filter(Boolean);
        }''')
        print("Table headers on CRM:", table_headers)

        rows = await page.evaluate('''() => {
            return Array.from(document.querySelectorAll('tbody tr')).slice(0, 3).map(el => el.innerText.replace(/\\n/g, ' | '));
        }''')
        print("First 3 rows on CRM:", rows)

        await browser.close()

if __name__ == '__main__':
    asyncio.run(inspect())
