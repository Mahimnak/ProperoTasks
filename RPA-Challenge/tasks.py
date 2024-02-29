from robocorp.tasks import task
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
from RPA.Robocorp.Vault import Vault
from RPA.Robocorp.WorkItems import WorkItem
from selenium.common.exceptions import NoSuchElementException, InvalidSelectorException

@task
def news_data():
    """Search for and scrape data from news websites."""
    details = fetch_workitems()
    
def navigate_webpage(details):
    """Navigate to the AP news website and collect data"""
    #initialise the driver using selenium and chromium web driver
    url = "https://apnews.com/"
    driver = webdriver.Chrome()
    #get the url of the website and open it in a browser window
    driver.get(url)
    
    #find the search button and click on it
    driver.find_element(By.XPATH,"//div[2]/bsp-header/div[1]/div[3]/bsp-search-overlay/button/svg[1]/use").click()
    #find the search input box and enter the phrase to be searched
    try:
        search_input = driver.find_element(By.XPATH, "//div[2]/bsp-header/div[1]/div[3]/bsp-search-overlay/div/form/label/input")
        if len(details) > 0:
            search_input.send_keys(details[0])
            driver.find_element(By.XPATH,"///div[2]/bsp-header/div[1]/div[3]/bsp-search-overlay/div/form/label/button/svg/use").click()
        else:
            print("Error: The details list is empty.")
    except NoSuchElementException:
        print("Error: Unable to locate the search input element.")
    except InvalidSelectorException:
        print("Error: The provided XPATH is invalid.")

    
    



def fetch_workitems():
    """Fetch the payload and details from the input work-items"""
    work_item = WorkItem()
    # Get the payload data from the work item
    payload = work_item.get_payload()
    
    # Access the data from the payload
    search_phrase = payload["search phrase"]
    category = payload["category"]
    number_of_months = payload["numberofmonths"]

    details = [search_phrase,category,number_of_months]

    return details

    
