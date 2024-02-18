from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
import openpyxl
from openpyxl import Workbook
import os
import sys
import shutil
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import zipfile
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
        zip_folder = find_and_add(current_emp_id,missing_docs_list)
        zip_missing_docs(zip_folder,current_emp_id)
        zip_filepath = "O:\\RPA\\Task 3 BotsDNA\\7-BackgroundVerification\\"+current_emp_id+"_MissingDocs"

        # file_input = driver.find_element(By.XPATH, '//*[@id="uploadedFile"]')
        # file_input.click()
        # time.sleep(5)


def zip_missing_docs(zip_folder,filename):
    zip_filename = filename+'_MissingDocs.zip'
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Walk through the directory and add files to the archive
        for root, dirs, files in os.walk(zip_folder):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, zip_folder))



def find_and_add(emp_id,missing_docs):
    def find_directory(directory_name, start_path):
        for root, dirs, files in os.walk(start_path):
            if directory_name in dirs:
                return os.path.join(root, directory_name)
        return None

    def find_file(file_name_start, start_path):
        for root, dirs, files in os.walk(start_path):
            for file in files:
                if file.startswith(file_name_start):
                    return os.path.join(root, file)
        return None
    def create_folder_and_copy_file(source_file_path, folder_name, destination_folder_path):
        if not destination_folder_path:
            directory_path = os.path.dirname(source_file_path)
        else:
            directory_path = destination_folder_path
        new_folder_path = os.path.join(directory_path, folder_name)
        os.makedirs(new_folder_path, exist_ok=True)
        new_file_path = os.path.join(new_folder_path, os.path.basename(source_file_path))
        shutil.copy2(source_file_path, new_file_path)
    
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

    
    start_path = 'O:\\RPA\\Task 3 BotsDNA\\7-BackgroundVerification\\Employee Documents\\Employee Documents'+'\\'+emp_id
 

    if start_path:
        for i in range(0, len(missing_docs)):
            file_name_start=missing_docs[i][:3]
            source_file_path = find_file(file_name_start, start_path)
            folder_name = str(emp_id)+"_MissingDocs"
            create_folder_and_copy_file(source_file_path,folder_name,start_path)
            zip_folder = start_path+"\\"+folder_name
        return zip_folder
    else:
        print(f"Directory not found: {emp_id}")
