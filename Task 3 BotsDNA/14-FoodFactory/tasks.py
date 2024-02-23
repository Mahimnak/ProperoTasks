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
from selenium.webdriver.support.relative_locator import locate_with
from selenium.webdriver.support.select import Select
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import pandas as pd


@task
def food_factory():
    """Collect details of various chips and create a report and send it to the respectiver person"""
    # Read the Excel file using pandas
    sheets = ["Chitoor","East Godavari","Guntur","Kadapa","Kurnool","Nellore","Prakasam","Vijayanagaram","Visakapatnam"]

    for i in range(0,len(sheets)):
        build_table(sheets[i])
        
def build_table(sheet_name):
    df = pd.read_excel("Andhra Pradesh.xlsx", sheet_name=sheet_name, engine='openpyxl')
    # Convert the dataframe to a list of lists
    data = df.values.tolist()
    nan_counter = 0

    # Remove all the occurrences of [nan, nan, nan, nan] from the data list and count them
    nan_counter = len(data) - len(list(filter(lambda x: not all(pd.isna(y) for y in x), data)))
    data = [row for row in data if not all(pd.isna(x) for x in row)]    
    number_locations = nan_counter+1
    
    # Create a PDF document
    doc = SimpleDocTemplate(sheet_name+".pdf", pagesize=letter)
    # Create a table with the data
    styles = getSampleStyleSheet()
    paragraph_style = ParagraphStyle(name="MyParagraph",parent=styles["BodyText"],lineSpacing=1.5)

    template = f"Janasanyog, Assam	Press Release No. 06 – Industries & Commerce \n PepsiCo India to invest Rs. 40000 crores in State \n Dispur, February 06: A delegation of PepsiCo India comprising Viraj Chouhan, Chief Government Affairs and Communication Officer-India Region; Rahul Sharma, Public Policy and Government Affairs Officer and Nitin Jindal, Associate Director, Business Planning called upon Industries and Commerce Minister Chandra Rajesh at Janata Bhawan, Dispur today. \n The team informed the Minister that PepsiCo India has submitted a proposal to set up a green field project for manufacturing Lays and Kurkure chips in Assam worth Rs. 40000 crores. They stated that the company would engage in contract farming (buy back policy) whereby it would provide seeds to the farmers and buy back the matured potatoes from them. The company would also provide handholding support to the farmers. Right now, PepsiCo India has been working with over 24,000 farmers in many cities in Andhra Pradesh.\n District Name: {sheet_name}\nTotal Locations: {number_locations}\n"
    paragraph = Paragraph(template, paragraph_style)

    table = Table(data[1:], colWidths=[100]*len(data[0]))
    # Apply a table style
    table.setStyle(TableStyle([
        ('FONT', (0, 0), (-1, -1), 'Helvetica', 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, 0), '#cccccc'),
        ('GRID', (0, 0), (-1, -1), 1, '#cccccc'),
    ]))

    doc.build([paragraph, table])