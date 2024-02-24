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
import io
from io import BytesIO
import pdfplumber
import PyPDF2

@task
def pay_bill():
    """Collect details from the pdfs and enter transaction details"""
    folder_path = "O:\\RPA\\Task 3 BotsDNA\\18-PayBill\\Payments"
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)
            extracted_text = extract_text_from_pdf(file_path)
            print(extracted_text.encode('utf-8'))

def extract_text_from_pdf(file_path):
    pdf_file_obj = open(file_path, 'rb')
    pdf_reader = PyPDF2.PdfReader(pdf_file_obj)
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text()
    pdf_file_obj.close()
    return text

    
