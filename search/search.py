import os
import search.constants as const
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import NoSuchElementException
from openpyxl import load_workbook
from datetime import datetime
import time

class Search:

    def __init__(self, driver_path = r"C:/SeleniumDrivers", teardown = False):

        self.driver_path = driver_path
        self.teardown = teardown

        os.environ['PATH'] += self.driver_path

        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        self.driver = webdriver.Chrome(options=options)
        self.driver.implicitly_wait(15)
        self.driver.maximize_window()

    def __enter__(self):

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        
        if self.teardown == True:
            
            self.driver.quit()

    def land_first_page(self):

        self.driver.get(const.BASE_URL)

        file_path = r"F:\5. Activities\fourbeats\search\excel.xlsx"
        
        wb = load_workbook(file_path)

        day = datetime.today().strftime("%A")
            
        sheet = wb[day]

        data = []

        count = 3

        self.driver.find_element(By.CSS_SELECTOR, "div[id='SIvCob'] a").click()

        for row in range(3, 13):
    
            cell_value = sheet[f'C{row}'].value

            data.append(cell_value)

        for i in data:

            temp = []

            self.driver.find_element(By.CSS_SELECTOR, "textarea").send_keys(i)

            time.sleep(2)

            values = self.driver.find_elements(By.CSS_SELECTOR, "li.sbct")

            for j in values:

                if j.text != "\n" and j.text != " " and j.text != "":

                    if j.is_displayed:

                        value = j.text.split("\n", 1)

                        temp.append(value[0])

            time.sleep(2)

            self.driver.find_element(By.CSS_SELECTOR, "textarea").clear()

            longest = max(temp, key=len)

            shortest = min(temp, key=len)

            sheet['D'+str(count)] = longest

            sheet['E'+str(count)] = shortest
            
            count+=1

        wb.save(file_path)

        self.driver.close()

