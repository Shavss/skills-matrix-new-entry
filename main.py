from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from pandas import ExcelFile
import pathlib
import time
import os

# Using environment variables for sensitive information
CHROMEDRIVER_PATH = os.getenv('CHROMEDRIVER_PATH', 'path/to/default/chromedriver')
COOKIES_PATH = os.getenv('COOKIES_PATH', str(pathlib.Path().absolute()) + '\\cookies')
TARGET_URL = os.getenv('TARGET_URL', 'https://example.com')  # Replace with default or empty string if appropriate

chrome_options = Options()
chrome_options.add_argument(f"user-data-dir={COOKIES_PATH}")
driver = webdriver.Chrome(CHROMEDRIVER_PATH, options=chrome_options)
driver.maximize_window()
driver.get(TARGET_URL)
driver.implicitly_wait(4)

s = 0
xls = ExcelFile("Skills Matrix_rapport data imput/Skills_Matrix/New_Skills_Matrix_Everyone.xlsx")
sn = xls.sheet_names[s]
ns = len(xls.sheet_names)
data = xls.parse(xls.sheet_names[s])
dic = data.to_dict()
id = sn.split(' ', 1)[0]

while s < ns:
    driver.get(f'{TARGET_URL}?hrid={id}')
    s = s + 1

    if s < ns:
        sn = xls.sheet_names[s]
        data = xls.parse(xls.sheet_names[s - 1])
        dic = data.to_dict()
        id = sn.split(' ', 1)[0]
    else:
        sn = "No more data"
        data = xls.parse(xls.sheet_names[s - 1])
        dic = data.to_dict()

    i = 1
    k = 1
    j = 0

    for l in range(95):
        print(s, ns, sn)
        print(i, k, j, l)

        try:
            new_skill = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//button[@title='Add New Skill']"))
            )
            new_skill.click()
        except Exception as e:
            print(f"Error finding 'Add New Skill' button: {e}")
            driver.quit()
            break

        WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it("expenseframe"))

        try:
            # Choosing Skill
            dropdown1 = driver.find_element(By.XPATH, "//div[@class='skillmatrix-add-new-content']//div[1]//select[1]")
            dd = Select(dropdown1)
            dd.select_by_index(i)

            # Choosing Skill Subcategory
            dropdown2 = driver.find_element(By.XPATH, "//div[@class='skillmatrix-add-new-container']//div[2]//select[1]")
            dd = Select(dropdown2)
            dd.select_by_index(k)

            time.sleep(2)

            # Selecting Competency Value
            k = k + 1
            c = dic["Competency_Value"][j]
            competency = driver.find_element(By.XPATH, f"//label[@for='newCompetency-{c}']")
            competency.click()

            # Selecting Experience Value
            e = dic["Experience_Value"][j]
            experience = driver.find_element(By.XPATH, f"//label[@for='newExperience-{e}']")
            experience.click()

            save_button = driver.find_element(By.XPATH, "//button[normalize-space()='Save']")
            save_button.click()

            j = j + 1
            driver.switch_to.parent_frame()
            time.sleep(2)

        except Exception as e:
            print(f"Error processing skill entry: {e}")
            cancel_button = driver.find_element(By.XPATH, "//button[normalize-space()='Cancel']")
            cancel_button.click()
            driver.switch_to.parent_frame()
            i = i + 1
            k = 1
            pass
