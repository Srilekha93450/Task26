# test_login_edge.py

import pytest
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from page_objects import LoginPage
from openpyxl import load_workbook
from datetime import datetime

# Path to the Microsoft Edge WebDriver executable
EDGE_DRIVER_PATH = r'C:\Users\WELCOME\Downloads\edgedriver_win64\msedgedriver.exe'

@pytest.fixture(scope='module')
def setup():
    # Use Edge WebDriver
    driver = webdriver.Edge(executable_path=EDGE_DRIVER_PATH)
    driver.maximize_window()
    yield driver
    driver.quit()

@pytest.mark.parametrize("test_id, username, password, date, time, tester", [
    (1, "Admin", "admin123", None, None, "Tester1"),
    (2, "Admin", "invalidpass", None, None, "Tester2"),
    (3, "invaliduser", "admin123", None, None, "Tester3"),
    (4, "Admin", "", None, None, "Tester4"),
    (5, "", "admin123", None, None, "Tester5")
])
def test_login(setup, test_id, username, password, date, time, tester):
    driver = setup
    driver.get("https://opensource-demo.orangehrmlive.com/")
    login_page = LoginPage(driver)
    
    login_page.enter_username(username)
    login_page.enter_password(password)
    login_page.click_login()

    # Explicit wait for login success or failure
    try:
        WebDriverWait(driver, 10).until(EC.url_to_be("https://opensource-demo.orangehrmlive.com/index.php/dashboard"))
        test_result = "Passed"
    except:
        test_result = "Failed"

    # Write result to Excel
    write_to_excel(test_id, username, password, date, time, tester, test_result)

def write_to_excel(test_id, username, password, date, time, tester, test_result):
    wb = load_workbook("test_data.xlsx")
    sheet = wb.active
    row = (test_id, username, password, date, time, tester, test_result)
    sheet.append(row)
    wb.save("test_data.xlsx")

    print(f"Test ID {test_id} completed with result: {test_result}")
