from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
import openpyxl
from openpyxl import Workbook
import os
import sys
import csv
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.support.relative_locator import locate_with
from selenium.webdriver.support.select import Select

@task
def background_check():
    """Background verification and handling missing documents"""
    url = "https://botsdna.com/BGV/"
    driver = webdriver.Chrome()
    driver.get(url)
    path = "O:\\RPA\\Task 3 BotsDNA\\7-BackgroundVerification"
    os.chdir(path)
    newfolder = ""
    # os.makedirs(newfolder)
    for i in range(0,1):
        current_emp_id = driver.find_element(By.XPATH, '//input[@id="CurrentEmpID"]').get_attribute("value")
        # Get the text from the input element
        missing_docs = driver.find_element(By.XPATH,"//*[@id='MissingDocs']").text
        missing_docs = re.sub(r'\s+', '', missing_docs)
        missing_docs_list = missing_docs.split('/')
        find_and_add(current_emp_id,missing_docs_list)

def find_and_add(emp_id,missing_docs):
    def find_directory(directory_name, start_path):
        for root, dirs, files in os.walk(start_path):
            if directory_name in dirs:
                return os.path.join(root, directory_name)
        return None

    def find_file(file_name_start, start_path):
        for root, dirs, files in os.walk(start_path):
            for file in files:
                if file_name_start in file:
                    return os.path.join(root, file)
        return None
    def open_directory(directory_path):
        if os.path.isdir(directory_path):
            if sys.platform == 'win32':
                os.startfile(directory_path)
            elif sys.platform == 'darwin':
                os.system('open ' + directory_path)
            else:
                os.system('xdg-open ' + directory_path)
        else:
            print(f"Directory not found: {directory_path}")

    
    start_path = 'O:\\RPA\\Task 3 BotsDNA\\7-BackgroundVerification\\Employee Documents\\Employee Documents'  
    target_directory_path = find_directory(emp_id, start_path)

    if target_directory_path:
        open_directory(target_directory_path)

    else:
        print(f"Directory not found: {emp_id}")
