import re
from playwright.sync_api import Playwright, sync_playwright, expect

def test() -> None:
 with sync_playwright() as playwright:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.cheapoair.ca/")
    page.get_by_role("tab", name="Packages").click()
    page.get_by_label("To where?").click()
    page.get_by_label("To where?").fill("bom")
    page.get_by_text("BOM - Mumbai, India").click()
    page.locator(".v2-date").first.click()
    page.get_by_label("Choose your Depart date.").click()
    page.get_by_label("27 December 2024").click()
    page.get_by_label("17 January").click()
    page.get_by_role("button", name="Room 2 Travelers").click()
    page.get_by_label("Added 2 ADULT").dblclick()
    page.get_by_label("Add Senior").click()
    page.get_by_role("button", name="Room 5 Travelers").click()
    page.get_by_role("button", name="Room 5 Travelers").click()
    page.get_by_role("button", name="Apply").click()
    page.get_by_role("button", name="Search Packages").click()
    page.get_by_role("button", name="Room 5 Travelers").click()
    page.get_by_label("Remove one remain 4 ADULT").click()
    page.get_by_text("I only need this hotel for part of my trip Select Hotel RatingAny").click()
    page.get_by_role("button", name="Search Packages").click()
    page.get_by_role("button", name="Got it").click()
    page.get_by_role("tab", name="Most Popular").click()
    page.get_by_label("Select Package with 5 Star Rating hotel at C$6030.07 Per person Flight+Hotel (").click()

    # ---------------------
    context.close()
    browser.close()

