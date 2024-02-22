from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
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
from selenium.webdriver.support.relative_locator import locate_with
from selenium.webdriver.support.select import Select
import json
@task
def vegetable_basket():
    """Get data from json file and update the details in the website"""
    cities = ["TodaysPrice\Hyderabad.json","TodaysPrice\Mangalore.json","TodaysPrice\Mysore.json","TodaysPrice\Nagpur.json","TodaysPrice\Visakhapatnam.json"]
    for i in range(0,len(cities)):
        get_data(cities[i])
def get_data(location):
    """Get data from json file"""
    with open(f"{location}", 'r') as file:
        data = json.load(file)

    vegitables = data['Vegitables']
    for veg in vegitables:
        code = veg['Code']
        Name = veg['Name']
        weight = veg['Weight']
        price = veg['Price']
        update_price(code,price)

def update_price(code,price):
    """Update price in the website"""
    url="https://botsdna.com/VegetableBasket/"
    driver = webdriver.Chrome()
    driver.get(url)
    driver.find_element(By.ID,"vegCode").send_keys(code)
    driver.find_element(By.ID,"Search").click()
    input_price = driver.find_element(By.ID,"Price")
    input_price.value = price
    driver.find_element(By.ID,"updateVeg").click()

        
