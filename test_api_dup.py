import asyncio
from playwright.async_api import async_playwright

async def check():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.goto('https://uat.inxify.com/login')
        await page.fill('input#username', 'quanhm2')
        await page.fill('input#password', '@Matkhau1')
        await page.click('button[type="submit"]')
        await page.wait_for_timeout(3000)
        
        await page.goto('https://uat.inxify.com/crm/guests')
        await page.wait_for_timeout(3000)
        
        # In Inxify, let's see what requests are made or inspect axios baseURL
        api_res = await page.evaluate('''async () => {
            // Find auth in localStorage
            let auth = '';
            for (let i = 0; i < localStorage.length; i++) {
                const k = localStorage.key(i);
                console.log('localStorage:', k);
            }
            // Fetch duplicates using window.fetch
            try {
                const r = await fetch('/guests/duplicates');
                return { status: r.status, text: await r.text() };
            } catch (err) {
                return { error: err.message };
            }
        }''')
        print("API Response:", api_res)
        await browser.close()

if __name__ == '__main__':
    asyncio.run(check())
