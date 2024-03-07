from robocorp.tasks import task
from workitems_data import fetch_workitems
from reuters import Reuters


@task
def news_data():
    """Search for and scrape data from news websites."""
    details = fetch_workitems()
    news = Reuters(details)
    news.create_excel()
    news.navigate_webpage()



    


    


























    #line 66 try:
    #initialise the driver using selenium and chromium web driver
    #from line 75 find the search button and click on it
    #     search_button = WebDriverWait(driver, 10).until(
    #         EC.element_to_be_clickable((By.CSS_SELECTOR, "svg[data-testid='SvgSearch']"))
    #     )
    #     search_button.click()
    #     time.sleep(5)
    #     #find the search input box and enter the phrase to be searched
    #     search_input = WebDriverWait(driver, 10).until(
    #         EC.element_to_be_clickable((By.CSS_SELECTOR,'input[data-testid="FormField:input"][type="search"][autocomplete="off"][spellcheck="false"][maxlength="256"][class="text__text__1FZLe text__dark-grey__3Ml43 text__regular__2N1Xr text__small__1kGq2 body__base__22dCE body__medium__2blzt form-field__input__7LFh3 form-field__default__1IHy7 search-bar__search-input__3ahqM"]'))
    #     )
    #     time.sleep(5)
    #     search_input.send_keys(details[0])
    #     time.sleep(5)
    #     search_submit_button = WebDriverWait(driver, 20).until(
    #         EC.element_to_be_clickable((By.CSS_SELECTOR, 'svg[data-testid="SvgSearch"][class="search-bar__icon__ORXTq search-bar__search-alt-icon__juWN_"]'))
    #     )
    #     time.sleep(5)
    #     search_submit_button.click()
    # except NoSuchElementException:
    #     print("Error: Unable to locate the search input element.")
    # except InvalidSelectorException:
    #     print("Error: The provided XPATH is invalid.")
    # except TimeoutException:
    #     print("Error: The element was not found within the specified time.")
    
    # try:
    #     #find the time span and click on it
    #     if details[2] == "1 month":
    #         past_month_option = WebDriverWait(driver, 10).until(
    #             EC.element_to_be_clickable((By.XPATH, "//*[@id='react-aria1233940870-:r1u:-option-Pastmonth']/span"))
    #         )
    #         past_month_option.click()
    #     elif details[2] == "Anytime":
    #         past_month_option = WebDriverWait(driver, 10).until(
    #             EC.element_to_be_clickable((By.XPATH, "//*[@id='react-aria1233940870-:r114:-option-Anytime']/span"))
    #         )
    #         past_month_option.click()
    #     elif details[2] == "24 hours":
    #         past_month_option = WebDriverWait(driver, 10).until(
    #             EC.element_to_be_clickable((By.XPATH, "//*[@id='react-aria1233940870-:r114:-option-Past24hours']/span"))
    #         )
    #         past_month_option.click()
    #     elif details[2] == "1 week":
    #         past_month_option = WebDriverWait(driver, 10).until(
    #             EC.element_to_be_clickable((By.XPATH, "//*[@id='react-aria1233940870-:r114:-option-Pastweek']/span"))
    #         )
    #         past_month_option.click()
    #     else:
    #         past_month_option = WebDriverWait(driver, 10).until(
    #             EC.element_to_be_clickable((By.XPATH, "//*[@id='react-aria1233940870-:r114:-option-Pastyear']/span"))
    #         )
    #         past_month_option.click()

    # except NoSuchElementException:
    #     print("Error: Unable to locate the date/time element.")
    # except InvalidSelectorException:
    #     print("Error: The provided XPATH is invalid.")
    # except TimeoutException:
    #     print("Error: The element was not found within the specified time.")

    #opening the worksheet