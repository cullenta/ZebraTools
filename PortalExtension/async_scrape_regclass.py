from playwright.async_api import async_playwright
import json
import asyncio

async def scrape_table(page):
    data = []

    while True:
        rows = await page.query_selector_all('tbody.MuiTableBody-root > tr')
        for row in rows:
            cells = await row.query_selector_all('th, td')
            row_data = [await cell.inner_text() for cell in cells]

            if len(row_data) != 9:
                continue

            data.append({
                "Day": row_data[0].strip(),
                "Times": row_data[1].strip(),
                "Stream": row_data[2].strip(),
                "Course": row_data[3].strip(),
                "Student ID": row_data[4].strip(),
                "Student Name": row_data[5].strip(),
                "Instructor Name": row_data[6].strip(),
                "Make Up Date": row_data[7].strip(),
                "Trial Date": row_data[8].strip()
            })

        next_button = page.locator('button[aria-label="Next Page"]')
        if await next_button.is_disabled():
            break
        await next_button.click()
        await page.wait_for_timeout(1000)

    return data

async def main_scraper():
    async with async_playwright() as p:
        browser = await p.webkit.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()
        await page.goto("https://portal.zebrarobotics.com/auth/login/")

        await page.fill('input[name="email"]', 'cullenta125@gmail.com')
        await page.fill('input[name="password"]', 'taite123')
        async with page.expect_navigation():
            await page.click('button[type="submit"]')

        await page.goto("https://portal.zebrarobotics.com/reporting/new-report/")
        await page.locator("form", has_text="Select...Select A Report").locator("svg").click()
        await page.click("#react-select-2-option-1-0")
        await page.get_by_role("checkbox").check()
        await page.get_by_role("button", name="Generate Report").click()
        await page.wait_for_selector('tbody.MuiTableBody-root > tr')

        table_data = await scrape_table(page)

        with open('./supabase_setup/report_data.json', 'w') as f:
            json.dump(table_data, f, indent=4)

        print(f"Scraped {len(table_data)} rows.")
        await browser.close()

        return table_data
