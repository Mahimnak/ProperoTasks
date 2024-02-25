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
import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
from docx.shared import Inches
from PIL import Image
import io
import requests
import zipfile

@task
def ip_ticketing():
    """Manage the tickets raised and fulfill and requirements"""
    driver = webdriver.Chrome()
    driver.get("https://botsdna.com/IPTicketingTool/index.html")

    temp=""
    ip_codes=[]
    ip_details=[]
    rows = driver.find_elements(By.XPATH,"//table/tbody/tr")
    folder_path="C:\\Users\\Mahimna\\Downloads"
    for i in range(0,1):
        rows[i].find_element(By.TAG_NAME,"i").click()
        time.sleep(2)
        ticket_det = rows[i].find_element(By.XPATH,f"//table/tbody/tr[{i+1}]/td[2]").text
        index_of_by = ticket_det.index("By")
        index_of_on = ticket_det.index("On")
        ticket_name = ticket_det[index_of_by + 3 : index_of_on - 1]
        ticket_number = ticket_det[7:11]
        zip_file_name = f"{ticket_number}.zip"
        text_file_name = f"{ticket_number}.txt"
        zip_file_path = find_zip_file(folder_path, zip_file_name)
        if zip_file_path:
            unzip_file(zip_file_path)
            lines = open_text_file(os.path.dirname(zip_file_path), text_file_name)
            for j in range(4,len(lines)-1):
                temp=lines[j]
                temp=temp[:-1]
                ip_codes.append(temp)
                temp=""
        else:
            print(f'Zip file {zip_file_name} not found in {folder_path}')
        for j in range(0,len(ip_codes)):
            ip_details.append(search_text_data(ip_codes[j]))
        
        send_message(ip_details,ticket_number,ip_codes[j], ticket_name)

def send_message(ip_details,t_num, ip_num, name):
    """Send the IP details and the messages to the required party"""
    driver2=webdriver.chrome()
    driver2.get("https://botsdna.com/IPTicketingTool/PeopleFinder.html")

    message = f"Dear {name}, \nPlease find below information of IPs, which you requested under ticket#{t_num} \n"
    temp=""
    for i in range(0,len(ip_details)):
        temp=f"IP:{ip_num}\n   Owner:{ip_details[i][0]}\n   Provider:{ip_details[i][1]}\n   ASN:{ip_details[i][2]}\n   City:{ip_details[i][3]}\n   Postal Code:{ip_details[i][4]}\n   Country:{ip_details[i][5]}\n   Coordinates:{ip_details[i][6]}\n"
        message+=temp
    
    search_box = driver2.find_element(By.ID,"SearchInput")
    search_box.send_keys(name)
    user = driver2.find_element(By.XPATH,"//div[2]/table/tbody/tr/td[1]/table/tbody/tr[2]/td/table/tbody/tr/td/b/span")
    user.click()
    message_area = driver2.find_element(By.XPATH,"//*[@id='msgBox']")
    message_area.send_keys(message)
    submit_button = driver2.find_element(By.ID,"SubmitButton")
    submit_button.click()

def search_text_data(ip_code):
    driver2=webdriver.chrome()
    driver2.get("https://botsdna.com/IPTicketingTool/IP%20Network.html")
    
    ip_details = driver2.find_element(By.XPATH,f'//*[@id="courts"]/tbody/tr[td[2][contains(text(), ip_code)]]')
    all_det =  ip_details.find_elements(By.TAG_NAME,"td")
    ip_owner = all_det[2].text
    provider =  all_det[3].text
    asn =  all_det[4].text

    driver2.get("https://botsdna.com/IPTicketingTool/IP%20Network.html")
    
    ip_details = driver2.find_element(By.XPATH,f'//*[@id="courts"]/tbody/tr[td[2][contains(text(), ip_code)]]')
    all_det =  ip_details.find_elements(By.TAG_NAME,"td")
    city = all_det[2].text
    postal_code =  all_det[3].text
    country =  all_det[4].text
    coordinates =  all_det[4].text
    final_details=[ip_owner,provider,asn,city,postal_code,country,coordinates]
    driver2.quit()
    return final_details

def find_zip_file(folder_path, zip_file_name):
    for root, dirs, files in os.walk(folder_path):
        if zip_file_name in files:
            return os.path.join(root, zip_file_name)
    return None

def unzip_file(zip_file_path):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(os.path.dirname(zip_file_path))

def open_text_file(folder_path, text_file_name):
    with open(os.path.join(folder_path, text_file_name), 'r') as file:
        lines = file.readlines()
    return lines

