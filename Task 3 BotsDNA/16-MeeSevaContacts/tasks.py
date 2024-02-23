from robocorp.tasks import task
from robocorp import browser
import openpyxl
from openpyxl import Workbook
import csv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.relative_locator import locate_with
from selenium.webdriver.support.select import Select

@task
def mee_seva():
    """Update Mee seva contact details"""
    driver = webdriver.Chrome()
    driver.get("https://botsdna.com/MeeSevaContacts/")
    # for i in range(1,14):
    #     for j in range(1,5):
    for i in range(1,14):
        for j in range(1,5):
            village = driver.find_element(By.XPATH,f"//table/tbody/tr[{i}]/td[{j}]/div/div/table/tbody/tr[2]/td[2]/input").get_attribute('value')
            driver2 = webdriver.Chrome()
            driver2.get("https://www.google.com/")
            m = driver2.find_element(By.NAME,"q")
            time.sleep(0.2)
            m.send_keys(village+" Meeseva Center Phone Number")
            time.sleep(0.2)
            #perform Google search with Keys.ENTER
            m.send_keys(Keys.ENTER)
            phone = driver.find_element(By.XPATH,"//*[@id='kp-wp-tab-overview']/div[2]/div/div/div/div[2]/div/div/span[2]/span/a/span").text
            primary_contact = driver.find_element(By.XPATH,f"//table/tbody/tr[{i}]/td[{j}]/div/div/table/tbody/tr[3]/td[2]/input").get_attribute('value')
            secondary_contact = driver.find_element(By.XPATH,f"/html/body/center/table/tbody/tr[{i}]/td[{j}]/div/div/table/tbody/tr[4]/td[2]/input").get_attribute('value')
            if primary_contact=="" or primary_contact!=phone:
                driver.find_element(By.XPATH,f"//table/tbody/tr[{i}]/td[{j}]/div/div/table/tbody/tr[3]/td[2]/input").send_keys(phone)
                driver.find_element(By.XPATH,f"//table/tbody/tr[{i}]/td[{j}]/div/div/table/tbody/tr[4]/td[2]/input").send_keys(primary_contact)
    driver.find_element(By.XPATH,"//*[@id='ContactSubmit']").click()
    
    