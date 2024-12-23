import re
from playwright.sync_api import Playwright, sync_playwright, expect

def test() -> None:
 with sync_playwright() as playwright:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.cheapoair.ca/")
    page.get_by_role("tab", name="Hotels").click()
    page.get_by_label("Destination", exact=True).click()
    page.get_by_label("Destination", exact=True).fill("bom")
    page.get_by_text("BOM - Mumbai Airport, India").click()
    page.get_by_label("16 January").click()
    page.get_by_label("31 January").click()
    page.get_by_role("button", name="Room 2 Travelers").click()
    page.get_by_label("Added 2 ADULT").click()
    page.get_by_label("Add children").click()
    page.get_by_label("7", exact=True).select_option("13")
    page.get_by_role("button", name="Search Hotels").click()
    page.locator("[data-test=\"dialog-button\"]").click()
    page.get_by_label("Sort by:").click()
    page.get_by_role("option", name="Price - Low to High").click()
    with page.expect_popup() as page1_info:
        page.get_by_label("Select The Bagicha residency").click()
    page1 = page1_info.value
    page1.get_by_label("Reserve a Room").click()
    page1.locator("[data-test=\"reserveSelectedRoom-details\"]").click()
    page1.locator("[data-test=\"link-key-things-to-know\"]").click()
    page1.locator("[data-test=\"close-icon-button\"]").click()

    # ---------------------
    context.close()
    browser.close()

