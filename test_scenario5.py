from playwright.sync_api import Playwright, sync_playwright
from openpyxl import Workbook

def scrape_flight_details_to_excel():
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # Navigate to the website and perform search
        page.goto("https://www.cheapoair.ca/")
        page.get_by_label("To where?").click()
        page.get_by_label("To where?").fill("bom")
        page.get_by_text("BOM - Mumbai, India").click()
        page.get_by_label("21 January").click()
        page.get_by_label("18 February").click()
        page.get_by_role("button", name="Search Flights").click()

        # Wait for the results page to load
        page.wait_for_selector("[data-test=\"flight-listing\"]")

        # Create a workbook for Excel
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Flight Details"
        sheet.append(["Airline", "Departure Time", "Arrival Time", "Price"])  # Add header row

        # Extract flight details
        flight_rows = page.query_selector_all("[data-test='flight-card']")  # Adjust selector if needed

        for row in flight_rows:
            try:
                airline = row.query_selector("[data-test='airline-name']").inner_text().strip()
                departure_time = row.query_selector("[data-test='departure-time']").inner_text().strip()
                arrival_time = row.query_selector("[data-test='arrival-time']").inner_text().strip()
                price = row.query_selector("[data-test='price']").inner_text().strip()

                # Append data to the Excel sheet
                sheet.append([airline, departure_time, arrival_time, price])
            except AttributeError:
                # Skip rows with missing data
                continue

        # Save the Excel file
        workbook.save("flight_details.xlsx")

        # Clean up
        context.close()
        browser.close()

if __name__ == "__main__":
    scrape_flight_details_to_excel()

