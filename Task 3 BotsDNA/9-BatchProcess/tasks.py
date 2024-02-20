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
from RPA.Email.ImapSmtp import ImapSmtpClient

@task
def batch_process():
    """Get batch transaction no. and send an email to the batch operator"""
    batches_url = "https://botsdna.com/BatchProcess/"
    driver1 = webdriver.Chrome()
    driver1.get(batches_url) 
    

    rows = driver1.find_elements(By.XPATH,"//table/tbody/tr")
    for row in rows:
        batch_details = row.text
        max_version = batch_details[7:9]
        min_version = batch_details[10:11]
        date = batch_details[20:22]+batch_details[18:20]+batch_details[14:18]
        order_type = batch_details[22:23]
        batch_code = batch_details[27:34]
        transaction_number = input_submit(max_version,min_version,date,order_type,batch_code)
        send_email(max_version,min_version,date,order_type,batch_code,transaction_number)

def input_submit(max_version, min_version, date, order_type, batch_code):
    """Input the batch details and get the transaction no."""
    submit_url = "https://botsdna.com/BatchProcess/SubmitBatch.html"
    driver2 = webdriver.Chrome()
    driver2.get(submit_url)
    batch_submit = driver2.find_element(By.XPATH,"//table/tbody/tr[1]/td[2]/input")
    batch_submit.send_keys(batch_code)
    max_version_input = driver2.find_element(By.XPATH,f"//table/tbody/tr[2]/td[2]/select/option[{max_version}]")
    max_version_input.click()
    min_version_input = driver2.find_element(By.XPATH,f"//table/tbody/tr[3]/td[2]/select/option[{min_version}]")
    min_version_input.click()
    date_input = driver2.find_element(By.XPATH,"//table/tbody/tr[4]/td[2]/input")
    date_input.send_keys(date)
    dubinium = driver2.find_element(By.XPATH,"//table/tbody/tr[5]/td[2]/input[1]")
    potassium = driver2.find_element(By.XPATH,"//table/tbody/tr[5]/td[2]/input[2]")
    if order_type=="P":
        potassium.click()
    else:
        dubinium.click()
    submit = driver2.find_element(By.XPATH,"//html/body/center/center/p/input")
    submit.click()
    #transaction number is displayed after submission - need to add delay here?
    transaction_number = driver2.find_element(By.ID,"TransNo").text
    return transaction_number

def send_email(max_version, min_version, date, order_type, batch_code,transaction_number):
    """Send an email to the batch operator"""
    client = ImapSmtpClient()

    # Connect to the email server
    client.smtp_connect(
        smtp_server="smtp.example.com",
        smtp_port=587,
        smtp_starttls=True,
        smtp_user="your_email@example.com",
        smtp_password="your_password"
    )

    # Send the email
    client.send_email(
        sender="your_email@example.com",
        recipients="mahimna77@gmail.com",
        subject="Submitted New Batch - Your transaction number is - "+transaction_number,
        body="Dear Team,\nNew batch has been submitted on"+date+"\nHere you can find the batch details,\nBatch Code:"+batch_code+"\n Max Version:"+max_version+"\n Min Version:"+min_version+"\n Batch Run Date:"+date+"\n Order type:"+order_type+"\n Transaction Number:"+transaction_number
    )

    # Disconnect from the email server
    client.smtp_disconnect()
