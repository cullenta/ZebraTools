from playwright.sync_api import sync_playwright
import json


def scrape_table(page):
    data = []
    
    while True:
        rows = page.query_selector_all('tbody.MuiTableBody-root > tr')
        for row in rows:
            cells = row.query_selector_all('th, td')
            row_data = [cell.inner_text().strip() for cell in cells]
            
            # Skip rows that don't have exactly 9 columns
            if len(row_data) != 9:
                continue

            data.append({
                "Day": row_data[0],
                "Times": row_data[1],
                "Stream": row_data[2],
                "Course": row_data[3],
                "Student ID": row_data[4],
                "Student Name": row_data[5],
                "Instructor Name": row_data[6],
                "Make Up Date": row_data[7],
                "Trial Date": row_data[8]
            })
    
        # Check if the "Next Page" button is disabled
        next_button = page.locator('button[aria-label="Next Page"]')
        if next_button.is_disabled():
            break
        else:
            next_button.click()
            page.wait_for_timeout(1000)  # Allow time for the next page to load

    return data

def main_scraper():
    with sync_playwright() as p:
        browser = p.webkit.launch(headless=False)
        context=browser.new_context()
        page = context.new_page()
        page.goto("https://portal.zebrarobotics.com/auth/login/")

        page.fill('input[name="email"]', 'cullenta125@gmail.com')
        page.fill('input[name="password"]', 'taite123')
        page.click('button[type="submit"]')

        with page.expect_navigation():
            page.click('button[type="submit"]')

        page.goto("https://portal.zebrarobotics.com/reporting/new-report/")

        print(page.title())

        page.locator("form", has_text="Select...Select A Report").locator("svg").click()

        # Select the option by ID
        page.click("#react-select-2-option-1-0")

        # Check the first visible checkbox (adjust if there are multiple)
        page.get_by_role("checkbox").check()

        # Click the "Generate Report" button by accessible name
        page.get_by_role("button", name="Generate Report").click()

        # Wait for the table to load
        page.wait_for_selector('tbody.MuiTableBody-root > tr')

        # Scrape the table
        table_data = scrape_table(page)

        # Save to JSON file
        with open('./supabase_setup/report_data.json', 'w') as f:
            json.dump(table_data, f, indent=4)

        print(f"Scraped {len(table_data)} rows.")
        browser.close()

        return table_data

