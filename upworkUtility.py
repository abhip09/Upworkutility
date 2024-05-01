from datetime import datetime
from os import name

import pandas as pd
from playwright.sync_api import sync_playwright

# Read configuration from Excel file
df = pd.read_excel("Input.xlsx", sheet_name='Configuration')
username = df.loc[0, 'Username']
password = df.loc[0, 'Password']
keywd = df.loc[0, 'Keyword']

# Define column names for DataFrame

column_names = ["Job Summary", "Description", "Skills", "About Client", "URL"]

# Create an empty DataFrame to store the scraped data
result_df = pd.DataFrame(columns=column_names)

skillsFilter = ["manual testing", "Automation", "API testing"]


def scrape_job_data(page):
    # Extract job data
    href = page.locator("//div//a[@data-test='slider-open-in-new-window UpLink']").get_attribute('href')
    URL = "https://www.upwork.com" + href
    title = page.locator("//h4[@class='d-flex align-items-center mt-0 mb-5']").text_content()
    description = page.locator("//div[@data-test='Description']").text_content()
    skills = page.locator("//section[@data-test='Expertise']").text_content()
    about_client = page.locator(
        "//div[@data-test='about-client-container AboutClientUserShared AboutClientUser']").text_content()

    # Append data to DataFrame
    result_df.loc[len(result_df)] = [title, description, skills, about_client, URL]
    print(result_df)

    #print("Test")



def run(playwright):
    global elements
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(record_video_dir="videos/")
    page = context.new_page()
    page.goto("https://www.upwork.com/ab/account-security/login")
    page.get_by_placeholder("Username or Email").click()
    page.get_by_placeholder("Username or Email").fill("sapanamorenagur@gmail.com")
    page.get_by_role("button", name="Continue", exact=True).click()
    page.get_by_role("textbox", name="Password").click()
    page.get_by_role("textbox", name="Password").fill("Sapana@26")
    page.get_by_role("button", name="Log in").click()
    page.goto("https://www.upwork.com/nx/find-work/")
    page.goto("https://www.upwork.com/nx/find-work/best-matches")
    page.get_by_placeholder("Username or Email").click()
    page.get_by_placeholder("Username or Email").fill("sapanamorenagur@gmail.com")
    page.get_by_role("button", name="Continue", exact=True).click()
    page.get_by_role("textbox", name="Password").click()
    page.get_by_role("textbox", name="Password").fill("Sapana@26")
    page.get_by_role("button", name="Log in").click()
    page.get_by_role("dialog").get_by_role("button", name="Close", exact=True).click()
    page.get_by_label("Close").click()
    page.get_by_placeholder("Search for jobs").click()
    page.get_by_placeholder("Search for jobs").fill("testing")
    page.wait_for_timeout(3000)
    page.keyboard.press('Enter')

    for i in skillsFilter:
        print(i)
        # page.locator("li:nth-child(2) > .air3-menu-item-text").click()
        page.get_by_role("button", name="Advanced Search").click()
        page.wait_for_timeout(3000)
        page.locator("[data-test=\"UpCAutoScroll\"]").get_by_role("searchbox").click()
        page.locator("[data-test=\"UpCTypeaheadScreenAdapter\"]").get_by_role("combobox").click()
        page.wait_for_timeout(3000)
        # page.locator("#typeahead-input-37").fill("Automation")
        # page.locator("[data-test=\"menu\"] span").first.click()
        page.locator("[data-test=\"UpCTypeaheadScreenAdapter\"]").get_by_role("combobox").fill(i)
        #  page.locator("#typeahead-input-37").fill(i)
        page.locator("[data-test=\"menu\"] span").first.click()
        # page.locator("#typeahead-input-37").fill("API Testing")
        # page.locator("[data-test=\"menu\"] span").first.click()
        page.wait_for_timeout(3000)
        page.locator("[data-test=\"submit-button\"]").click()
        page.wait_for_timeout(3000)
        page.click("//div[@data-test='jobs_per_page UpCDropdown']")
        page.wait_for_timeout(3000)
        page.keyboard.press('ArrowDown')
        page.keyboard.press('ArrowDown')
        page.keyboard.press('Enter')
        page.wait_for_timeout(5000)

        # Extract job data
        elements = page.locator(
            "//article//div//small//span[contains(text(),'minutes ago')] | //article//div//small//span[contains(text(),"
            "'hours ago')] | //article//div//small//span[contains(text(),'hour ago')]").element_handles()

        for element in elements:
            element.click()
            scrape_job_data(page)
            page.keyboard.press('Escape')  # Close the job detail popup
        page.locator("//button[contains(text(), 'Clear filters')]").click()
        page.wait_for_timeout(3000)

    context.close()
    browser.close()
    path = page.video.path()
    print(path)
with sync_playwright() as playwright:
    run(playwright)


current_time = datetime.now()
sheet_name = current_time.strftime('Sheet_%Y-%m-%d_%H-%M-%S')
print(result_df)
# Write DataFrame to Excel file
# Path to your Excel file
file_path = "Input.xlsx"

# Attempt to write to a new sheet in the existing workbook
try:
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='new') as writer:
        result_df.to_excel(writer, sheet_name=sheet_name, index=False)
except FileNotFoundError:
    # If the file does not exist, create it with the DataFrame
    result_df.to_excel(file_path, sheet_name=sheet_name, index=False)
