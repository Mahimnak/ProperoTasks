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
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

@task
def work_anniversary():
    """send a mail wishing the employee a happy work anniversary"""
    driver = webdriver.Chrome()
    driver.get("https://botsdna.com/WorkAnniversary/")

    workbook = openpyxl.load_workbook('WorkAnniversary.xlsx')
    sheet = workbook.active

    rows = driver.find_elements(By.XPATH,"//table/tbody/tr")
    for i in range(1,len(rows)):
        details = rows[i].find_elements(By.TAG_NAME,"td")
        emp_id = details[0].text
        emp_name = details[1].text
        man_id = details[2].text
        doj = details[3].text
        doj = doj[-4:]
        experience = 2024-int(doj)
        for row_num, cell in enumerate(sheet["A"], start=1):
            if cell.value == emp_id:
                row_value = row_num
                break
        emp_email = sheet.cell(row = row_value, column = 3).value
        send_email(emp_email,experience, emp_name)
        driver.get("https://botsdna.com/WorkAnniversary/YearsOfExp.html")
        full_name = emp_name.split()
        driver.find_element(By.ID,"surname").send_keys(full_name[-1])
        driver.find_element(By.ID,"name").send_keys(full_name[0])
        driver.find_element(By.ID,"years_of_experience").send_keys(experience)
        driver.find_element(By.ID,"manager_id").send_keys(man_id)
        driver.find_element(By.XPATH,"//*[@id='submission-form']/input[5]").click()


def send_email(e_email, exp, emp_name):
    """Sending an email to the employee"""
    sender = 'sender@example.com'
    recipients = e_email
    subject = f'{emp_name}, Happy Anniversary'

    # Set the email content
    body_html = f'''
    <html><body>
    <p>This certificate is awarded to</p>
    <p><strong>{emp_name}</strong></p>
    <p>for [his/her] outstanding service, tireless effort, and constant support</p>
    <p>for the Propero and its projects for the last {exp} years.</p>
    </body></html>
    '''

    # Set the email message
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = subject

    # Attach the email body as HTML content
    msg.attach(MIMEText(body_html, 'html'))

    # Send the email
    server = smtplib.SMTP('smtp.example.com', 587)
    server.starttls()
    server.login(sender, 'password')
    server.sendmail(sender, recipients, msg.as_string())
    server.quit()
