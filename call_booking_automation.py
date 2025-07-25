# Importing required packages
import datetime
import calendar
import os
import time
from autohubble import hubble_query_to_df, PRESTO
import pandas as pd
import numpy as np
import csv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

################################################################################################ 
# Hubble query to call booking data

sql_data = '''
    with cte as (
        select row_number() over (partition by region, email, country order by created desc) as row_num,
		    active,region, email, country, created    
	    from usertables.ansuman_call_booking_automation
    )select region, email, country 
    from cte where row_num = 1 and active = true
    '''

# df_sql_data = hubble_query_to_df(sql_data, PRESTO, force_refresh=True)
df_sql_data = hubble_query_to_df(sql_data, PRESTO)
print(df_sql_data)

# Dataframe to store slot counts
results_df = pd.DataFrame(columns=['Region', 'Country', 'Email', 'SlotCount', 'CheckedAt'])

################################################################################################ 
#Selenium function

# Initialize Chrome WebDriver
service = Service()
driver = webdriver.Chrome(service=service)

def check_slots_for_region(driver,region, email, country):    
    try:
        driver.get("https://stripe.com/in/contact/sales")
        time.sleep(3) # wait to open

        ########################################################## First form filling
        email_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//input[@name='email']"))) # Find email field 
        email_field.clear() # Clear any place holder data
        email_field.send_keys(email) # Fill new data
        
        country_dropdown = driver.find_element(By.XPATH, "//select[@name='Country__c']") # Find country dropdown 
        country_dropdown.click() # Click to open the drop down
        
        country_option = driver.find_element(By.XPATH, f"//option[@data-country-name='{country}']") # Select the country
        country_option.click() # And click on it
                
        continue_button = driver.find_element(By.XPATH, "//button[contains(@class, 'ContactSalesFormStep__continueButton')]")
        continue_button.click()
        
        time.sleep(1) # wait to open

        ########################################################## First form filling
        first_name_field = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//input[@name='firstName']")))# Fill first name
        first_name_field.send_keys("test")
        
        last_name_field = driver.find_element(By.XPATH, "//input[@name='lastName']") # Fill last name
        last_name_field.send_keys("test")
        
        phone_field = driver.find_element(By.XPATH, "//input[@name='phone']") # Fill phone number
        phone_field.send_keys("9999999990")
        
        website_field = driver.find_element(By.XPATH, "//input[@name='website']") # Fill company website
        website_field.send_keys("test.com")
        
        job_level_dropdown = driver.find_element(By.XPATH, "//select[@name='Job_Level_Form__c']") # Find Job level dropdown 
        job_level_dropdown.click()
        job_level_option = driver.find_element(By.XPATH, "//select[@name='Job_Level_Form__c']/option[@value='C-level']") # Select job level - C-level
        job_level_option.click()
        
        job_function_dropdown = driver.find_element(By.XPATH, "//select[@name='Job_Function_Form__c']") # Find Job function dropdown 
        job_function_dropdown.click()
        job_function_option = driver.find_element(By.XPATH, "//select[@name='Job_Function_Form__c']/option[@value='Marketing']")# Select job function - Marketing
        job_function_option.click()

        submit_button = driver.find_element(By.XPATH, "//button[@data-test-id='csf-contact-information-submit-button']")
        submit_button.click()
        
        print("Waiting for calendar page to load...")
        time.sleep(20)  # Wait for page to load
        print("Wait time completed")

         # Method 1: Look for buttons with current date in title
        now = datetime.datetime.now()
        day_of_week = calendar.day_name[now.weekday()]  # e.g., "Thursday"
        month_name = calendar.month_name[now.month]     # e.g., "July"
        day_number = now.day                            # e.g., 24

        print(f"Looking for buttons with date pattern: {day_of_week}, {month_name} {day_number}")
        date_pattern = f"{day_of_week}, {month_name} {day_number}"

        iframe_leandatabookit = driver.find_element(By.ID, "LeanDataBookitFrame")
        print("Found LeanDataBookitFrame by ID")
        driver.switch_to.frame(iframe_leandatabookit)

        iframe_bookit = driver.find_element(By.CLASS_NAME, "bookit-frame")
        print("Found bookit-frame by ID")
        driver.switch_to.frame(iframe_bookit)

        confirm_buttons = driver.find_elements(By.XPATH, "//button[contains(@class, 'timeslot-button')]/div[contains(@class, 'confirm-button-div')]/button[contains(@class, 'confirm-button')]")
        print(f"Found {len(confirm_buttons)} confirm buttons")

        if len(confirm_buttons) > 0:
            slot_count = len(confirm_buttons)
        
        time.sleep(2)  # Wait for page to load
        
        return slot_count        
        
    except Exception as e:
        print(f"Error in form submission or checking time slots: {str(e)}")
        return 0

try:
    # Process each region in the dataframe
    for index, row in df_sql_data.iterrows():
        region = row['region']
        email = row['email']
        country = row['country']
        
        print(f"Checking slots for {region}, {country}...")
        slot_count = check_slots_for_region(driver, region, email, country)
        
        # Add results to the dataframe
        results_df = pd.concat([results_df, pd.DataFrame([{
            'Region': region,
            'Country': country,
            'Email': email,
            'SlotCount': slot_count,
            'CheckedAt': datetime.datetime.now()
        }])], ignore_index=True)
        
        # Add a small delay between requests to prevent getting blocked
        time.sleep(2)
    
    script_directory = os.path.dirname(__file__)
    result_csv_path = os.path.join(script_directory, 'result_data.csv')
    results_df.to_csv(result_csv_path, index=False)
    print(f"\nResults saved to {result_csv_path}")
        
finally:
    # Only close the browser after all checks are done
    driver.quit()
