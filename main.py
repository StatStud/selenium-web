from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time

chrome_options = Options()
## uncomment this block out if you want to run headless version (without browser)
# chrome_options.add_argument("--no-sandbox")
# chrome_options.add_argument("--headless=new")
# chrome_options.add_argument("--disable-dev-shm-usage")
# chrome_options.add_argument("--enable-logging")
# chrome_options.add_argument("--v=1")

download_dir = "~/Downloads"
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 10)

version = driver.capabilities['browserVersion']
print(f"Chrome version: {version}")

def login(driver):
    driver.get("example.com")
    username = os.environ.get('username')
    password = os.environ.get('password')
    username_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='username']")))
    username_field.send_keys(username)
    password_field = driver.find_element(By.XPATH, "//input[@name='password']")
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)

def nav_to_workbook(driver):
    item_name_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@data-testid='itemName']")))
    item_name_element.click()
    wait.until(lambda d: len(d.window_handles) > 1)
    window_handles = driver.window_handles
    driver.switch_to.window(window_handles[-1])
    worksheet_container_element = wait.until(EC.element_to_be_clickable((By.ID, "BADFE346-3BF6-4848-AE28-5841A8D87A04")))
    worksheet_container_element.click()

def access_date_select(driver):
    try:
        time_range_label = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@class='mb5' and @data-testid='dateTimeRangeLabel']")))
        if time_range_label:
            return
    except:
        pass
    try:
        export_excel_element = wait.until(EC.element_to_be_clickable((By.ID, "export-excel")))
        export_excel_element.click()
    except:
        try:
            export_button = wait.until(EC.element_to_be_clickable((By.ID, "import-export")))
            export_button.click()
            export_excel_element = wait.until(EC.element_to_be_clickable((By.ID, "export-excel")))
            export_excel_element.click()
        except:
            try:
                driver.refresh()
                export_excel_element = wait.until(EC.element_to_be_clickable((By.ID, "export-excel")))
                export_excel_element.click()
            except:
                driver.refresh()
                export_excel_element = wait.until(EC.element_to_be_clickable((By.ID, "import-export")))
                export_excel_element.click()
                export_excel_element = wait.until(EC.element_to_be_clickable((By.ID, "export-excel")))
                export_excel_element.click()

def reload_page(driver):
    driver.refresh()
    time.sleep(20)

def prepare_input_field(driver):
    num_back_spaces = 40
    actions = ActionChains(driver)
    actions.send_keys(Keys.ARROW_RIGHT)
    for _ in range(num_back_spaces):
        actions.send_keys(Keys.BACK_SPACE)
    actions.perform()

def export_data(driver, start_date, end_date):
    range_start_container = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='nowrap dateTimeEntry specRange specRangeStart']//p[@id='rangeStart' and @data-testid='valueInput_rangeStart']")))
    range_start_container.click()
    range_start_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@id='rangeStart' and @data-testid='editableInput_rangeStart']")))
    prepare_input_field(driver)
    range_start_input.send_keys(start_date)
    range_end_container = driver.find_element(By.XPATH, "//div[@class='nowrap dateTimeEntry specRange specRangeEnd']//p[@id='rangeEnd' and @data-testid='valueInput_rangeEnd']")
    range_end_container.click()
    range_end_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@id='rangeEnd' and @data-testid='editableInput_rangeEnd']")))
    prepare_input_field(driver)
    range_end_input.send_keys(end_date)
    value_with_units_input = driver.find_element(By.XPATH, "//input[@data-testid='valueWithUnitsInput']")
    value_with_units_input.click()
    value_with_units_input.send_keys(Keys.ARROW_RIGHT * 2)
    value_with_units_input.send_keys(Keys.BACK_SPACE * 2)
    value_with_units_input.send_keys("1")
    selected_value_div = wait.until(EC.presence_of_element_located((By.XPATH, "(//div[@class='specOpenSelect css-w54w9q-singleValue'])[2]")))
    selected_value = selected_value_div.text.strip()
    if selected_value != "second(s)":
        dropdown = driver.find_element(By.XPATH, "//div[@data-testid='valueWithUnitsUnits_filter']")
        dropdown.click()
        actions = ActionChains(driver)
        actions.send_keys(Keys.ARROW_UP * {"minute(s)": 1, "hour(s)": 2, "days(s)": 3}.get(selected_value, 0))
        actions.send_keys(Keys.ENTER)
        actions.perform()
    checkbox = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@data-testid='gridOriginEnabled-input']")))
    if checkbox.get_attribute("value").lower() == "false":
        checkbox.click()
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    export_button = wait.until(EC.element_to_be_clickable((By.ID, "exportExcelButton")))
    export_button.click()
    wait_minutes = 15
    wait_long = WebDriverWait(driver, 60*wait_minutes)
    wait_long.until(EC.staleness_of(export_button))
    time.sleep(10)

if __name__ == "__main__":
    login(driver)
    nav_to_workbook(driver)
    dates_lst = [
        '1/1/2022 12:00 AM',
        '2/1/2022 12:00 AM',
        '3/1/2022 12:00 AM',
        '4/1/2022 12:00 AM',
        '5/1/2022 12:00 AM',
        '6/1/2022 12:00 AM',
        '7/1/2022 12:00 AM',
        '8/1/2022 12:00 AM',
        '9/1/2022 12:00 AM',
        '10/1/2022 12:00 AM',
        '11/1/2022 12:00 AM',
        '12/1/2022 12:00 AM',
        '1/1/2023 12:00 AM'
    ]
    for i in range(len(dates_lst)-1):
        start_idx = dates_lst[i]
        end_idx = dates_lst[i+1]
        print(f"starting with range {start_idx}, to {end_idx}")
        try:
            access_date_select(driver)
        except:
            reload_page(driver)
            access_date_select(driver)
        print("\t we are now downloading data...")
        try:
            start = time.time()
            export_data(driver, start_date=start_idx, end_date=end_idx)
            end = time.time()
        except:
            reload_page(driver)
            start = time.time()
            export_data(driver, start_date=start_idx, end_date=end_idx)
            end = time.time()
        print(f"\t finished in {end-start:.2f} sec")
    print("done!")
