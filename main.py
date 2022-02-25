from selenium import webdriver
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from fill_forms import FillForms

CHROME_DRIVE_PATH = "C:\DRIVERS\chromedriver_win32\chromedriver"

service = Service(CHROME_DRIVE_PATH)
driver = webdriver.Chrome(service=service)
driver.get("https://www.rightmove.co.uk/")
driver.implicitly_wait(10)

# Deal with cookies
try:
    driver.find_element(By.XPATH, "//button[normalize-space()='Allow all cookies']").click()
except ElementClickInterceptedException:
    driver.find_element(By.XPATH, "//button[@title='Allow all cookies']").click()
else:
    driver.find_element(By.NAME, "typeAheadInputField").click()

# Select Chesterfield in Derbyshire
driver.find_element(By.NAME, "typeAheadInputField").click()
actions = ActionChains(driver)
actions.send_keys("Chesterfield, Derbyshire")
actions.send_keys(Keys.ENTER)
actions.perform()

# Select radius 10 miles
driver.find_element(By.ID, "radius").click()
actions = ActionChains(driver)
actions.send_keys(Keys.DOWN*6)
actions.send_keys(Keys.ENTER)
actions.perform()

# Select minimum 3 bedrooms
driver.find_element(By.ID, "minBedrooms").click()
actions = ActionChains(driver)
actions.send_keys(Keys.DOWN*4)
actions.send_keys(Keys.ENTER)
actions.perform()

# Select Property Type: Houses
driver.find_element(By.ID, "displayPropertyType").click()
actions = ActionChains(driver)
actions.send_keys(Keys.DOWN)
actions.send_keys(Keys.ENTER)
actions.perform()

# Select Added to site in last 24h
driver.find_element(By.ID, "maxDaysSinceAdded").click()
actions = ActionChains(driver)
actions.send_keys(Keys.DOWN)
actions.send_keys(Keys.ENTER)
actions.perform()

# Submit
driver.find_element(By.ID, "submit").click()
number_of_pages = int(driver.find_element(By.XPATH, '//span[@data-bind="text: total"]').text)

# Scrape required data
all_locations = []
all_prices = []
all_links = []

for n in range(number_of_pages):

    # 1. Find and append locations
    locations = driver.find_elements(By.XPATH, "//meta[@itemprop='streetAddress']")
    for location in locations:
        get_location = location.get_attribute("content")
        all_locations.append(get_location)

    # 2. Find and append prices
    prices = driver.find_elements(By.CLASS_NAME, "propertyCard-priceValue")
    for price in prices:
        price_text = price.text
        all_prices.append(price_text)

    # 3. Find and append links
    links = driver.find_elements(By.XPATH, "//div[@class='l-searchResult is-list']//descendant::a[@class='propertyCard-link']")
    for link in links:
        link_href = link.get_attribute("href")
        all_links.append(link_href)

    # Deal with stale element exception after moving to the next page
    try:
        def find(driver):
            next_btn = driver.find_element(By.XPATH, "//button[@title='Next page']")
            if next_btn:
                return next_btn
            else:
                return False
        next_button = WebDriverWait(driver, 5).until(find)
        next_button.click()
    except ElementClickInterceptedException:
        break
driver.close()

# print(all_locations)
# print(all_prices)
# print(all_links)


fill_form = FillForms(all_locations, all_prices, all_links)

# Save all_locations, all_prices, all_link to excel spreadsheet
fill_form.fill_excel()

# Save all_locations, all_prices, all_link to google drive excel spreadsheet
fill_form.fill_google_spreadsheet()

# Save all_locations, all_prices, all_link to csv file
fill_form.fill_csv()
