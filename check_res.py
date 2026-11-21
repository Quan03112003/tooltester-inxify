import asyncio
from playwright.async_api import async_playwright

async def check_reservations():
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
        
        # Search for any warning or note
        res_info = await page.evaluate('''() => {
            const rows = Array.from(document.querySelectorAll('tbody tr')).map(tr => tr.innerText.replace(/\\n/g, ' | '));
            const tags = Array.from(document.querySelectorAll('.ant-tag, span[class*=\"tag\"]')).map(t => t.innerText.trim());
            return { rows: rows.slice(0, 10), tags: tags };
        }''')
        print("Rows:", res_info['rows'][:5])
        print("Tags:", res_info['tags'][:20])
        
        await browser.close()

if __name__ == '__main__':
    asyncio.run(check_reservations())
