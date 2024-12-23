import re
from playwright.sync_api import Playwright, sync_playwright, expect

def test() -> None:
  with sync_playwright() as playwright:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.cheapoair.ca/")
    page.get_by_label("To where?").click()
    page.get_by_label("To where?").fill("bom")
    page.get_by_text("BOM - Mumbai, India").click()
    page.get_by_label("Choose a departure date.").click()
    page.get_by_label("14 January").click()
    page.get_by_label("Choose a return date. expanded").click()
    page.get_by_label("18 February").click()
    page.get_by_role("button", name="Search Flights").click()
    page.goto("https://www.cheapoair.ca/air/listing?&d1=YEA&r1=BOM&dt1=01/14/2025&d2=BOM&r2=YEA&dt2=02/18/2025&tripType=ROUNDTRIP&cl=ECONOMY&ad=1&se=0&ch=0&infs=0&infl=0&flightsearch=true")

    # ---------------------
    context.close()
    browser.close()
