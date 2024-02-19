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
import pdfrw


@task
def points_of_contact():
    """Find the details from the web and create pdf files for each contact"""
    url="https://botsdna.com/poc/"
    driver = webdriver.Chrome()
    driver.get(url)
    # Wait until page is loaded completely
    wait = WebDriverWait(driver, 2)

    rows = driver.find_elements(By.XPATH,"//table/tbody/tr[]")
    for row in rows:
        project_code = row.find_element(By.XPATH,"//tbody/tr[]/td[1]").text
        dev_phone = row.find_element(By.XPATH,"//tbody/tr[]/td[2]").text
        man_phone = row.find_element(By.XPATH,"//tbody/tr[]/td[3]").text
        if dev_phone and man_phone:
            poc = Select(driver.find_element(By.XPATH,"//*[@id='ContactType']"))
            poc.select_by_visible_text("Both")
        elif dev_phone:
            poc = Select(driver.find_element(By.XPATH,"//*[@id='ContactType']"))
            poc.select_by_visible_text("Developer")
        elif man_phone:
            poc = Select(driver.find_element(By.XPATH,"//*[@id='ContactType']"))
            poc.select_by_visible_text("Manager")
        description(project_code, dev_phone,man_phone)

def description(proj_code, developer, manager):
    """Find the project description file for the specific code and save it to a pdf file"""
    def find_file(file_name_start, start_path):
        for root, dirs, files in os.walk(start_path):
            for file in files:
                if file.startswith(file_name_start):
                    path = os.path.join(root, file)
                    with open(path, 'r') as file:
                    # Read the content of the file
                        content = file.read()
                        return content
        return None
    def create_pdf(path, content):
        pdf = pdfrw.PdfReader("template.pdf")

    # Write some text
        page = pdf.pages[0]
        page.Annots.append(pdfrw.annot.create_annot(page, "FreeText", pdfrw.Rect(100, 750, 100, 100), "Developer No."+developer))
        page.Annots.append(pdfrw.annot.create_annot(page, "FreeText", pdfrw.Rect(100, 750, 100, 100), "Manager No."+manager))
        page.Annots.append(pdfrw.annot.create_annot(page, "FreeText", pdfrw.Rect(100, 750, 100, 100), "Project Description:"+content))
    # Save the PDF file in a specific location
        pdfrw.PdfWriter().write(path, pdf.pages)

    start_path = "O:\\RPA\\Task 3 BotsDNA\\8-PointsofContact\\Required Files\\Project Description"
    file_content = find_file(proj_code, start_path)
    path2 = "O:\\RPA\\Task 3 BotsDNA\\8-PointsofContact\\Required Files"
    create_pdf(path2,file_content)

