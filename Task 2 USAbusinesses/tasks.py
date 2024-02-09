from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from robocorp import vault
from RPA.Excel.Application import Application
import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import Workbook
import csv
from lxml import etree 
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import time

@task
def scrape_business_details():
    """Scrape the details of USA based companies in all the provided sectors."""
    urls_industry = ["http://www.gbguides.com/advertising/","http://www.gbguides.com/appliances/","http://www.gbguides.com/arts/","http://www.gbguides.com/computers/","http://www.gbguides.com/construction/","http://www.gbguides.com/education/","http://www.gbguides.com/employment/","http://www.gbguides.com/entertainment/","http://www.gbguides.com/events/","http://www.gbguides.com/finance/","http://www.gbguides.com/health/","http://www.gbguides.com/houses/","http://www.gbguides.com/professionals/","http://www.gbguides.com/professionals/","http://www.gbguides.com/shopping/","http://www.gbguides.com/travel/"]
    switch_industry(urls_industry)

def switch_industry(urls):
    """Go from one industry to the other, for example: advertising, events, travel, etc."""
    
    get_data(urls[0])

def get_data(current_url):
    """Get data from a specific page and write it into an Excel file."""
    wb = Workbook()
    ws = wb.active

    driver = webdriver.Chrome()
    for i in range(2201, 5546):
        if i != 1:
            next_url = f"{current_url}{i}/"
        else:
            next_url = current_url
        
        driver.get(next_url)
        resultados_div = driver.find_element(By.CLASS_NAME, "resultados")
        li_elements = resultados_div.find_elements(By.TAG_NAME, "li")
        
        for li in li_elements:
            try:
                company_name = li.find_element(By.TAG_NAME, 'a').text
            except NoSuchElementException:
                company_name = ""
            try:
                location = li.find_element(By.CLASS_NAME, 'lista-direccion').text
            except NoSuchElementException:
                location = ""
            try:
                phone = li.find_element(By.CLASS_NAME, 'telnumero').text
            except NoSuchElementException:
                phone = ""
            try:
                website_element = li.find_element(By.CLASS_NAME, 'lista-web').find_element(By.TAG_NAME, 'a')
                website = website_element.text
            except NoSuchElementException:
                website = ""
            
            # Append data to the worksheet
            ws.append([company_name, location, phone, website])

        if i % 100 == 0:
            # Save the workbook every 100 iterations
            wb.save(f'data_{i}.xlsx')
            # Clear the worksheet for the next batch of data
            ws = wb.active
            ws.delete_rows(1, ws.max_row)
    
    # Save the final workbook
    wb.save('data_final.xlsx')

    # Close the WebDriver session
    driver.quit()
