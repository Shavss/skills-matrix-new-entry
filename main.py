from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

from pandas import *
import pathlib
import time
#import credentials


PATH = "C:\Program Files (x86)\chromedriver.exe"
script_directory = pathlib.Path().absolute()

chrome_options = Options()
chrome_options.add_argument(f"user-data-dir={script_directory}\\cookies")
driver = webdriver.Chrome(PATH, options=chrome_options)
chrome_options.add_argument(f"user-data-dir={script_directory}\\cookies")
driver.maximize_window()
driver.get("https://davidchipperfield.rapport3.com/applications/hr/hr_14.asp")
driver.implicitly_wait(4)

#We used cookies in order to log us in into the system so the below part is not necessary
# (
    #username = credentials.TEST_IO_USERNAME
    #password = credentials.TEST_IO_PASSWORD
    #username_input = driver.find_element(By.ID, 'i0116')
    #username_input.send_keys(username)
    #username_input.send_keys(Keys.RETURN)
    #time.sleep(5)
    #password_input = driver.find_element(By.ID, 'i0118')
    #password_input.send_keys(password)
    #password_input.send_keys(Keys.RETURN)
    #time.sleep(10)
    #next_button = driver.find_element(By.XPATH, "//input[@type='submit']")
    #next_button.click()
# )



s = 0
xls = ExcelFile("Skills Matrix_rapport data imput/Skills_Matrix/New_Skills_Matrix_Everyone.xlsx")
sn = xls.sheet_names[s]
ns = len(xls.sheet_names)
data = xls.parse(xls.sheet_names[s])
dic = data.to_dict()
id = sn.split(' ', 1)[0]

while s < ns:

    driver.get(f'https://davidchipperfield.rapport3.com/applications/hr/hr_14.asp?hrid={id}')
    s = s + 1

    # Need to acknowledge what if there is no more sheets after the last one
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

    for l in range(95):  # Maybe calculate amount of data rows and + the amount of errors at the end

        print(s, ns, sn)
        print(i, k, j, l)

        try:
            new_skill = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//button[@title='Add New Skill']"))
            )
            new_skill.click()

        except:
            driver.quit()

        WebDriverWait(driver, 3).until(EC.frame_to_be_available_and_switch_to_it("expenseframe"))

        try:
            # Choosing Skill
            class DemoDropdownSingleSelect():
                def demo_dropdown(self):
                    dropdown1 = driver.find_element(By.XPATH,
                                                    "//div[@class='skillmatrix-add-new-content']//div[1]//select[1]")
                    dd = Select(dropdown1)
                    dd.select_by_index(i)


            dddemo = DemoDropdownSingleSelect()
            dddemo.demo_dropdown()


            # Choosing Skill Subcategory
            class DemoDropdownSingleSelect():
                def demo_dropdown(self):
                    dropdown2 = driver.find_element(By.XPATH,
                                                        "//div[@class='skillmatrix-add-new-container']//div[2]//select[1]")
                    dd = Select(dropdown2)
                    dd.select_by_index(k)


            dddemo2 = DemoDropdownSingleSelect()
            dddemo2.demo_dropdown()

            time.sleep(2)

            # Selecting Competency Value
            k = k + 1
            c = dic["Competency_Value"][j]
            if c == 1:
                 competency = driver.find_element(By.XPATH, "//label[@for='newCompetency-1']")
            elif c == 2:
                competency = driver.find_element(By.XPATH, "//label[@for='newCompetency-2']")
            elif c == 3:
                competency = driver.find_element(By.XPATH, "//label[@for='newCompetency-3']")
            elif c == 4:
                competency = driver.find_element(By.XPATH, "//label[@for='newCompetency-4']")
            elif c == 5:
                competency = driver.find_element(By.XPATH, "//label[@for='newCompetency-5']")

            competency.click()

            # Selecting Experience Value
            e = dic["Experience_Value"][j]
            if e == 0:
                experience = driver.find_element(By.XPATH, "//label[@for='newExperience-0']")
            elif e == 1:
                experience = driver.find_element(By.XPATH, "//label[@for='newExperience-1']")
            elif e == 2:
                experience = driver.find_element(By.XPATH, "//label[@for='newExperience-2']")
            elif e == 3:
                experience = driver.find_element(By.XPATH, "//label[@for='newExperience-3']")
            elif e == 4:
                experience = driver.find_element(By.XPATH, "//label[@for='newExperience-4']")
            elif e == 5:
                experience = driver.find_element(By.XPATH, "//label[@for='newExperience-5']")

            experience.click()

            save_button = driver.find_element(By.XPATH, "//button[normalize-space()='Save']")
            save_button.click()

            j = j + 1

            driver.switch_to.parent_frame()

            time.sleep(2)
        except:
            cancel_button = driver.find_element(By.XPATH, "//button[normalize-space()='Cancel']")
            cancel_button.click()
            driver.switch_to.parent_frame()
            i = i + 1
            k = 1
            pass






