import re
from playwright.sync_api import Playwright, sync_playwright
from openpyxl import Workbook

def test_extract_flight_data_to_excel() -> None:
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # Navigate to the flight search page
        page.goto("https://www.cheapoair.ca/")
        page.get_by_label("To where?").click()
        page.get_by_label("To where?").fill("bom")
        page.get_by_text("BOM - Mumbai, India").click()
        page.get_by_label("Choose a departure date.").click()
        page.get_by_text("One Way", exact=True).click()
        page.get_by_label("Choose a departure date.").click()
        page.get_by_label("13 January").click()
        page.get_by_role("button", name="Traveler").click()
        page.get_by_label("Add falseseniors").click()
        page.get_by_label("Added 1 falseadults").click()
        page.get_by_role("button", name="Search Flights").click()

        # Wait for the search results to load (update selector as needed)
        page.wait_for_selector(".flight-results")  # Replace with the appropriate selector

        # Extract airlines and pricing details
        airline_elements = page.query_selector_all(".airline-name")  # Update with the selector for airline names
        price_elements = page.query_selector_all(".price-display")  # Update with the selector for prices

        # Collect the data
        flight_data = []
        for airline, price in zip(airline_elements, price_elements):
            airline_name = airline.inner_text().strip()
            flight_price = price.inner_text().strip()
            flight_data.append({"Airline": airline_name, "Price": flight_price})

        # Save data to Excel
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Flight Data"
        sheet.append(["Airline", "Price"])  # Add header row

        for flight in flight_data:
            sheet.append([flight["Airline"], flight["Price"]])

        # Save the Excel file
        workbook.save("flight_data.xlsx")

        # Clean up
        context.close()
        browser.close()

if __name__ == "__main__":
    test_extract_flight_data_to_excel()
