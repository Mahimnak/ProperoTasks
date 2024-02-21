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
import base64
import hashlib
from Crypto.Cipher import AES
import sys
import binascii

@task
def remote_desktop():
    """Connect to remote desktop and check schedule"""
    driver = webdriver.Chrome()
    url="https://botsdna.com/RemoteDesktop/"
    driver.get(url)
    ip_address = driver.find_elements(By.XPATH,"//*[@id='ipTable']/tbody/tr/td[1]") 
    user_id = driver.find_elements(By.XPATH,"//*[@id='ipTable']/tbody/tr/td[2]") 
    encrypted = driver.find_elements(By.XPATH,"//*[@id='ipTable']/tbody/tr/td[3]/i") 

    for i in range(1,25):
        driver.get(url)
        current_ip = ip_address[i].text
        current_user = user_id[i].text
        onclick = encrypted[i].get_attribute("onclick")
        encrypted_password = onclick.split("'")[1]
        decrypted_password = decrypt(encrypted_password,AES.MODE_GCM)
        driver.find_element(By.ID,"Computer").send_keys(current_ip)
        driver.find_element(By.ID,"UserName").send_keys(current_user)
        driver.find_element(By.ID,"Password").send_keys(decrypted_password)
        driver.find_element(By.XPATH,"//*[@id='RemoteConnect']").click()
        driver.find_element(By.XPATH,"//*[@id='SignOff'']").click()

        
def decrypt(ciphertext, mode):
    key="botsDNA"
    (ciphertext,  authTag, nonce) = ciphertext
    encobj = AES.new(key,  mode, nonce)
    return(encobj.decrypt_and_verify(ciphertext, authTag))
    
