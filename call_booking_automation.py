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

# sql_data = '''
#     with cte as (
#         select row_number() over (partition by region, email, country order by created desc) as row_num,
# 		    active,region, email, country, created    
# 	    from usertables.ansuman_call_booking_automation
#     )select region, email, country 
#     from cte where row_num = 1 and active = true
#     limit 1
#         '''

# # df_sql_data = hubble_query_to_df(sql_data, PRESTO, force_refresh=True)
# df_sql_data = hubble_query_to_df(sql_data, PRESTO)
# print(df_sql_data)

script_directory = os.path.dirname(__file__)
latest_file_path = os.path.join(script_directory, "test_case.csv")

data = pd.read_csv(latest_file_path, header=0, encoding='ISO-8859-1') # Read the test_case file
column_data = data[["region","email","country"]] # Validate column names and extract data

# Ensure no duplicates in the mapping
original_count = len(column_data)
column_data = column_data.drop_duplicates()
if original_count > len(column_data):
    print(f"Removed {original_count - len(column_data)} duplicate entries from the file")

print(f"Loaded {len(column_data)} test cases from {latest_file_path}") # Final data loaded
df_sql_data = column_data

# Dataframe to store slot counts
results_df = pd.DataFrame(columns=['Region', 'Country', 'Email', 'Day', 'Date', 'SlotCount', 'CheckedAt'])

# Date
now = datetime.datetime.now()
day_of_week = calendar.day_name[now.weekday()]  # e.g., "Thursday"
month_name = calendar.month_name[now.month]     # e.g., "July"
year = now.year      # e.g., "2025"
day_number = now.day         
print(f"Looking for buttons with date pattern: {month_name} {year}") #August 2025
print(f"Day:{day_number}") # 1

################################################################################################ 
#Selenium function

# Initialize Chrome WebDriver
service = Service()
driver = webdriver.Chrome(service=service)

def check_slots_for_region(driver,region, email, country):    
    try:
        driver.get("https://stripe.com/in/contact/sales")
        time.sleep(3) # wait to open
        region_results = []

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
        time.sleep(15)  # Wait for page to load
        print("Wait time completed")

        iframe_leandatabookit = driver.find_element(By.ID, "LeanDataBookitFrame")
        print("Found LeanDataBookitFrame by ID")
        driver.switch_to.frame(iframe_leandatabookit)

        iframe_bookit = driver.find_element(By.CLASS_NAME, "bookit-frame")
        print("Found bookit-frame by ID")
        driver.switch_to.frame(iframe_bookit)

        ########################################################### Counting slots code

        print("Starting to process day")
        # Get all available day buttons
        available_days = []
        try:
            all_day_buttons = driver.find_elements(By.XPATH, "//button[contains(@class, 'day-btn')]")
            print(f"Found {len(all_day_buttons)} total day buttons")
            
            for day_btn in all_day_buttons:
                try:
                    day_num = day_btn.text.strip()
                    day_title = day_btn.get_attribute("title")
                    is_selected = "selected-date" in day_btn.get_attribute("class")
                    
                    # Only add buttons with numeric values
                    if day_num.isdigit():
                        available_days.append({
                            'button': day_btn,
                            'day_num': int(day_num),  # Convert to int for proper sorting
                            'title': day_title,
                            'is_selected': is_selected
                        })
                        print(f"Day {day_num}: title={day_title}, selected={is_selected}")
                except Exception as e:
                    print(f"Error processing a day button: {str(e)}")
            
            print(f"Successfully processed {len(available_days)} day buttons")
            
            # Sort the days by their numeric value to process in ascending order
            available_days.sort(key=lambda d: d['day_num'])
            print("Sorted days for processing:")
            for day in available_days:
                print(f"  Day {day['day_num']}: {day['title']}")
                
        except Exception as e:
            print(f"Error finding day buttons: {str(e)}")
        
        # Process up to 8 days
        days_processed = 0
        days_tobe_processed = 8
        
        for day in available_days:
            if days_processed >= days_tobe_processed:
                break
                
            day_button = day['button']
            day_title = day['title']
            day_num = day['day_num']
            is_selected = day['is_selected']
            
            # Click on the day button (skip if it's the first day and already selected)
            if not is_selected or days_processed > 0:
                print(f"Clicking on day {day_num} - {day_title}")
                try:
                    day_button.click()
                    time.sleep(0.5)  # Wait for slots to update
                except Exception as e:
                    print(f"Error clicking on day {day_num}: {str(e)}")
                    days_processed += 1
                    continue
            else:
                print(f"Day {day_num} is already selected")
            
            # Extract date from the title (e.g., "Friday, August 1")
            date_part = day_title
            
            # Count the slots for this day using the date in the title attribute
            try:
                # Find slots that have this date in their title
                timeslot_buttons = driver.find_elements(By.XPATH, f"//button[contains(@class, 'timeslot-button') and contains(@title, '{date_part} at')]")
                slot_count = len(timeslot_buttons)/2
                print(f"Found {slot_count} slots for day {day_num} ({day_title})")
                
                # Store the results
                region_results.append({
                    'Region': region,
                    'Country': country,
                    'Email': email,
                    'Day': days_processed + 1,
                    'Date': day_title,
                    'SlotCount': slot_count,
                    'CheckedAt': datetime.datetime.now()
                })
            except Exception as e:
                print(f"Error counting slots for day {day_num}: {str(e)}")
                
            days_processed += 1
            time.sleep(0.5)  # Add delay between day clicks
        
        print(f"Completed processing {days_processed} days")
        
        # Switch back to default content
        driver.switch_to.default_content()
        return region_results
        
    except Exception as e:
        print(f"Error in form submission or checking time slots: {str(e)}")
        return region_results

try:
    # Process each region in the dataframe
    for index, row in df_sql_data.iterrows():
        region = row['region']
        email = row['email']
        country = row['country']

        print(f"Checking slots for {region}, {country}...")
        slot_count = check_slots_for_region(driver, region, email, country)
        
        # Add results to the dataframe in the new format
        if slot_count:  # Check if slot_count has data
            # Create a base row with region, country, email
            row_data = {
                'Region': region,
                'Country': country,
                'Email': email
            }
            
            # Add each date as a column with its slot count
            for day_data in slot_count:
                date_key = day_data['Date']  # Use the date as column name
                row_data[date_key] = day_data['SlotCount']
            
            # Add the row to results dataframe
            results_df = pd.concat([results_df, pd.DataFrame([row_data])], ignore_index=True)
            
        else:
            # If no slots found, add with basic info only
            row_data = {
                'Region': region,
                'Country': country,
                'Email': email
            }
            results_df = pd.concat([results_df, pd.DataFrame([row_data])], ignore_index=True)
        
        time.sleep(2)  # small delay between requests to prevent getting blocked
    
    # Fill NaN values with 0 (for dates that don't have data for some regions)
    results_df = results_df.fillna(0)
    
    # Define base columns that should appear first
    base_columns = ['Region', 'Country', 'Email']
    
    # Identify date columns (exclude unwanted columns)
    unwanted_columns = ['CheckedAt', 'Date', 'Day', 'SlotCount', 'Day_Title', 'No_Data']
    date_columns = [col for col in results_df.columns if col not in base_columns and col not in unwanted_columns]
    
    # Sort date columns in ascending chronological order
    if date_columns:
        try:
            # Parse dates and sort them chronologically
            date_columns_with_parsed = []
            for date_col in date_columns:
                try:
                    # Extract date from formats like "Friday, August 15" or "2024-01-15"
                    # Try to parse the date string
                    parsed_date = pd.to_datetime(date_col, errors='coerce')
                    if pd.isna(parsed_date):
                        # Try alternative parsing for "Day, Month DD" format
                        import re
                        # Extract month and day from strings like "Friday, August 15"
                        match = re.search(r'(\w+),\s*(\w+)\s+(\d+)', date_col)
                        if match:
                            month_day = f"{match.group(2)} {match.group(3)}"
                            parsed_date = pd.to_datetime(f"2024 {month_day}", errors='coerce')
                    
                    if pd.notna(parsed_date):
                        date_columns_with_parsed.append((parsed_date, date_col))
                    else:
                        # If all parsing fails, use current date as fallback for sorting
                        date_columns_with_parsed.append((pd.Timestamp.now(), date_col))
                except:
                    date_columns_with_parsed.append((pd.Timestamp.now(), date_col))
            
            # Sort by parsed date and extract column names
            date_columns_with_parsed.sort(key=lambda x: x[0])
            date_columns = [col[1] for col in date_columns_with_parsed]
            
        except Exception as e:
            print(f"Date sorting failed: {e}, using alphabetical sort")
            date_columns.sort()
    
    # Create final column order: base columns + sorted date columns only
    final_columns = base_columns + date_columns
    
    # Select only the desired columns
    results_df = results_df[final_columns]
    
    result_csv_path = os.path.join(script_directory, 'result_data.csv')
    results_df.to_csv(result_csv_path, index=False)
    print(f"\nResults saved to {result_csv_path}")
    print("Columns in final output:", results_df.columns.tolist())
    
finally: 
    driver.quit()
