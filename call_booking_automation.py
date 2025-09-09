# Importing required packages
import datetime
import calendar
import os
import time
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
# Function 1: Load Data (test cases and previous results)
def load_data():
    print("\n=== FUNCTION: LOADING DATA ===")
    script_directory = os.path.dirname(__file__)
    
    # Load test cases
    latest_file_path = os.path.join(script_directory, "test_case.csv")
    data = pd.read_csv(latest_file_path, header=0, encoding='ISO-8859-1')
    column_data = data[["region", "email", "country"]]
    
    # Ensure no duplicates in the mapping
    original_count = len(column_data)
    column_data = column_data.drop_duplicates()
    if original_count > len(column_data):
        print(f"Removed {original_count - len(column_data)} duplicate entries from the file")
    
    print(f"Loaded {len(column_data)} test cases from {latest_file_path}")
    
    # Load previous results if available
    result_csv_path = os.path.join(script_directory, "result_data.csv")
    previous_results = None
    if os.path.exists(result_csv_path):
        try:
            previous_results = pd.read_csv(result_csv_path)
            print(f"Loaded previous results from {result_csv_path} for delta comparison")
        except Exception as e:
            print(f"Error loading previous results: {e}")
    else:
        print("No previous results found. This will be the baseline run.")
    
    return column_data, previous_results, result_csv_path

################################################################################################
# Function 2: Check Slots for Region
def check_slots_for_region(driver, region, email, country):
    print(f"\n=== FUNCTION: CHECKING SLOTS for {region}, {country} ===")
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
        
        time.sleep(2) # wait to open

        ########################################################## Check for revenue form
        try:
            # Check if the revenue form appears
            revenue_form = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//div[@data-js-controller='Revenue']"))
            )
            print("Revenue form detected, analyzing available options...")
            
            # Wait for the form to fully render
            time.sleep(2)
            
            # Find the visible revenue options container
            visible_container_xpath = "//div[@data-js-controller='Revenue']//div[@data-js-controller='ButtonRadio']/div[contains(@class, 'ContactSalesButtonRadio') and not(contains(@class, 'hidden'))]"
            visible_container = driver.find_element(By.XPATH, visible_container_xpath)
            
            if visible_container:
                # Find all the option labels in the visible container
                option_labels = visible_container.find_elements(By.XPATH, "./label[contains(@class, 'ContactSalesButtonRadioOption')]")
                
                print(f"Found {len(option_labels)} revenue options in the visible container")
                
                # Print all options for debugging
                for idx, label in enumerate(option_labels):
                    try:
                        text_element = label.find_element(By.XPATH, ".//div[contains(@class, 'ContactSalesButtonRadioOption__text')]")
                        text = text_element.text.strip()
                        print(f"Option {idx+1}: Text='{text}'")
                    except:
                        print(f"Option {idx+1}: Unable to read text")
                
                if len(option_labels) >= 2:
                    # Select the second to last option
                    target_index = len(option_labels) - 2
                    target_label = option_labels[target_index]
                    
                    try:
                        target_text = target_label.find_element(By.XPATH, ".//div[contains(@class, 'ContactSalesButtonRadioOption__text')]").text.strip()
                        print(f"Selecting option {target_index+1}: '{target_text}'")
                    except:
                        print(f"Selecting option {target_index+1}")
                    
                    # Click the label directly (this worked in the previous attempt)
                    try:
                        target_label.click()
                        print("Successfully clicked the option label")
                    except Exception as e:
                        print(f"Failed to click option label: {e}")
                    
                    # Wait for the option to be selected and for any UI changes
                    time.sleep(2)
                    
                    # Look for a continue button that might appear after selection
                    # The continue button may have appeared or become enabled after selection
                    continue_button_xpath = "//div[@data-js-controller='Revenue']//button[contains(text(), 'Continue') or contains(@class, 'continueButton')]"
                    
                    try:
                        # Wait for the continue button to appear/be enabled
                        continue_button = WebDriverWait(driver, 3).until(
                            EC.element_to_be_clickable((By.XPATH, continue_button_xpath))
                        )
                        print("Continue button found and is clickable")
                        continue_button.click()
                        print("Clicked continue button")
                    except Exception as e:
                        print(f"Continue button not found or not clickable: {e}")
                        # Just continue with the flow, as the form might auto-advance
                        print("Proceeding to the next form step anyway")
                else:
                    print(f"Not enough options found, only {len(option_labels)} options")
            else:
                print("No visible revenue options container found")
                
        except Exception as e:
            # Revenue form did not appear or couldn't be handled
            print(f"No revenue form detected or couldn't be handled: {e}")
        
        ########################################################## Second form filling
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

        # Find the iframe and switch to it
        iframe_leandatabookit = driver.find_element(By.ID, "LeanDataBookitFrame")
        print("Found LeanDataBookitFrame by ID")
        driver.switch_to.frame(iframe_leandatabookit)

        iframe_bookit = driver.find_element(By.CLASS_NAME, "bookit-frame")
        print("Found bookit-frame by class name")
        driver.switch_to.frame(iframe_bookit)

        # Get today's date
        today = datetime.datetime.now().date()
        target_end_date = today + datetime.timedelta(days=7)  # Next 8 days (including today)
        print(f"Looking for slots from {today} to {target_end_date}")
        
        # Function to parse date from day title
        def parse_day_title(title):
            try:
                import re
                match = re.search(r'(\w+),\s*(\w+)\s+(\d+)', title)
                if match:
                    month = match.group(2)
                    day = int(match.group(3))
                    
                    month_num = {
                        'January': 1, 'February': 2, 'March': 3, 'April': 4, 
                        'May': 5, 'June': 6, 'July': 7, 'August': 8, 
                        'September': 9, 'October': 10, 'November': 11, 'December': 12
                    }.get(month)
                    
                    year = today.year
                    # Adjust year if month is earlier than current month
                    if month_num < today.month:
                        year += 1
                    
                    return datetime.date(year, month_num, day)
            except:
                pass
            return None
        
        # Process current month
        processed_days = 0
        processed_dates = set()
        
        # Get the current month shown
        current_month = None
        try:
            month_selector = driver.find_element(By.XPATH, "//p-dropdown[@id='month-selector']//span[contains(@class, 'p-dropdown-label')]")
            current_month = month_selector.text
            print(f"Current month in calendar: {current_month}")
        except:
            print("Couldn't identify current month")
        
        # Process current month days
        day_buttons = driver.find_elements(By.XPATH, "//button[contains(@class, 'day-btn')]")
        print(f"Found {len(day_buttons)} day buttons in current month")
        
        for day_btn in day_buttons:
            try:
                day_num = day_btn.text.strip()
                if not day_num.isdigit():
                    continue
                    
                day_title = day_btn.get_attribute("title")
                day_date = parse_day_title(day_title)
                
                if day_date and today <= day_date <= target_end_date:
                    # Click on the day to see available slots
                    day_btn.click()
                    time.sleep(0.5)
                    
                    # Count slots for this day
                    timeslot_buttons = driver.find_elements(By.XPATH, f"//button[contains(@class, 'timeslot-button') and contains(@title, '{day_title} at')]")
                    slot_count = len(timeslot_buttons)/2
                    print(f"Found {slot_count} slots for {day_date} ({day_title})")
                    
                    region_results.append({
                        'Region': region,
                        'Country': country,
                        'Email': email,
                        'Day': (day_date - today).days + 1,
                        'Date': day_title,
                        'SlotCount': slot_count,
                        'CheckedAt': datetime.datetime.now()
                    })
                    
                    processed_dates.add(day_date)
                    processed_days += 1
            except Exception as e:
                print(f"Error processing day: {e}")
        
        # If we need more days and haven't processed all 8 days yet, check next month
        if len(processed_dates) < 8:
            try:
                print("Need more days, trying to navigate to next month...")
                
                # Click month selector
                month_selector = driver.find_element(By.XPATH, "//p-dropdown[@id='month-selector']")
                month_selector.click()
                time.sleep(1)
                
                # Try to select next month (second option)
                next_month_options = driver.find_elements(By.XPATH, "//p-dropdownitem")
                if len(next_month_options) > 1:
                    next_month = next_month_options[1].find_element(By.XPATH, "./li")
                    next_month_text = next_month.text
                    print(f"Found next month option: {next_month_text}")
                    
                    # Only click if it's a different month
                    if next_month_text != current_month:
                        next_month.click()
                        time.sleep(2)
                        
                        # Process days in next month
                        day_buttons = driver.find_elements(By.XPATH, "//button[contains(@class, 'day-btn')]")
                        print(f"Found {len(day_buttons)} day buttons in next month")
                        
                        for day_btn in day_buttons:
                            # Stop if we've already processed 8 days
                            if len(processed_dates) >= 8:
                                break
                                
                            try:
                                day_num = day_btn.text.strip()
                                if not day_num.isdigit():
                                    continue
                                    
                                day_title = day_btn.get_attribute("title")
                                day_date = parse_day_title(day_title)
                                
                                if day_date and today <= day_date <= target_end_date and day_date not in processed_dates:
                                    # Click on the day to see available slots
                                    day_btn.click()
                                    time.sleep(0.5)
                                    
                                    # Count slots for this day
                                    timeslot_buttons = driver.find_elements(By.XPATH, f"//button[contains(@class, 'timeslot-button') and contains(@title, '{day_title} at')]")
                                    slot_count = len(timeslot_buttons)/2
                                    print(f"Found {slot_count} slots for {day_date} ({day_title})")
                                    
                                    region_results.append({
                                        'Region': region,
                                        'Country': country,
                                        'Email': email,
                                        'Day': (day_date - today).days + 1,
                                        'Date': day_title,
                                        'SlotCount': slot_count,
                                        'CheckedAt': datetime.datetime.now()
                                    })
                                    
                                    processed_dates.add(day_date)
                            except Exception as e:
                                print(f"Error processing day in next month: {e}")
            except Exception as e:
                print(f"Error navigating to next month: {e}")
        
        print(f"Total days processed: {len(region_results)}")
        
        # Switch back to default content
        driver.switch_to.default_content()
        return region_results
        
    except Exception as e:
        print(f"Error in form submission or checking time slots: {str(e)}")
        return []

# Function 3: Process and format results (continued)
def process_results(slot_data_list, df_sql_data):
    print("\n=== FUNCTION: PROCESSING RESULTS ===")
    
    # Initialize results dataframe
    results_df = pd.DataFrame(columns=['Region', 'Country', 'Email'])
    
    # Process each region's slot data
    for index, row in df_sql_data.iterrows():
        region = row['region']
        email = row['email']
        country = row['country']
        
        # Find the slot data for this region
        region_slot_data = None
        for data in slot_data_list:
            if data and len(data) > 0:
                if data[0]['Region'] == region:
                    region_slot_data = data
                    break
        
        # Add results to the dataframe in the wide format
        if region_slot_data:  # Check if slot data exists for this region
            # Create a base row with region, country, email
            row_data = {
                'Region': region,
                'Country': country,
                'Email': email
            }
            
            # Add each date as a column with its slot count
            for day_data in region_slot_data:
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
                except Exception as e:
                    print(f"Error parsing date {date_col}: {str(e)}")
                    date_columns_with_parsed.append((pd.Timestamp.now(), date_col))
            
            # Sort by parsed date and extract column names
            date_columns_with_parsed.sort(key=lambda x: x[0])
            date_columns = [col[1] for col in date_columns_with_parsed]
            
        except Exception as e:
            print(f"Date sorting failed: {e}, using alphabetical sort")
            date_columns.sort()
    
    # Create final column order: base columns + sorted date columns
    final_columns = base_columns + date_columns
    
    # Select only the desired columns
    results_df = results_df[final_columns]
    
    print(f"Processed results with {len(results_df)} regions and {len(date_columns)} dates")
    return results_df, date_columns

################################################################################################
# Function 4: Compare data and calculate deltas
def calculate_deltas(current_results, previous_results, date_columns):
    print("\n=== FUNCTION: CALCULATING DELTAS ===")
    
    if previous_results is None or previous_results.empty:
        print("No valid previous results to compare with. Skipping delta calculation.")
        return None
    
    try:
        # Create a copy of current results for delta calculations
        delta_df = current_results.copy()
        
        # Collect all date columns from both current and previous results
        all_date_columns = set(date_columns)
        
        # Add date columns from previous results that might not be in current results
        prev_date_cols = [col for col in previous_results.columns 
                          if col not in ['Region', 'Country', 'Email'] 
                          and not col.endswith('(Δ)')]
        all_date_columns.update(prev_date_cols)
        
        # Sort date columns (similar to the process in process_results function)
        try:
            # Parse dates and sort them chronologically
            date_columns_with_parsed = []
            for date_col in all_date_columns:
                try:
                    # Extract date from formats like "Friday, August 15" or "2024-01-15"
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
                except Exception as e:
                    print(f"Error parsing date {date_col}: {str(e)}")
                    date_columns_with_parsed.append((pd.Timestamp.now(), date_col))
            
            # Sort by parsed date and extract column names
            date_columns_with_parsed.sort(key=lambda x: x[0])
            sorted_date_columns = [col[1] for col in date_columns_with_parsed]
            
        except Exception as e:
            print(f"Date sorting failed: {e}, using alphabetical sort")
            sorted_date_columns = sorted(list(all_date_columns))
        
        # Add missing date columns to delta_df with zeros
        for date_col in sorted_date_columns:
            if date_col not in delta_df.columns:
                delta_df[date_col] = 0
                print(f"Added missing date column '{date_col}' with zeros")
                
        # For each date column, calculate and add a delta column
        delta_columns_added = 0
        
        # Create a final list of columns for ordering
        final_columns = ['Region', 'Country', 'Email']
        
        for date_col in sorted_date_columns:
            # Add the current value column to final columns
            if date_col in delta_df.columns:
                final_columns.append(date_col)
            
            # Create delta column name
            delta_col = f"{date_col} (Δ)"
            
            # Calculate delta values
            delta_values = []
            
            for _, curr_row in delta_df.iterrows():
                key = (curr_row['Region'], curr_row['Country'], curr_row['Email'])
                curr_val = float(curr_row.get(date_col, 0))
                
                # Find matching row in previous results
                prev_val = 0  # Default if not found
                for _, prev_row in previous_results.iterrows():
                    if (prev_row['Region'], prev_row['Country'], prev_row['Email']) == key:
                        prev_val = float(prev_row.get(date_col, 0))
                        break
                
                delta = curr_val - prev_val
                delta_values.append(delta)
            
            # Add delta column to dataframe
            delta_df[delta_col] = delta_values
            final_columns.append(delta_col)
            delta_columns_added += 1
            
        # Reorder columns to ensure values and deltas are paired
        delta_df = delta_df[final_columns]
        
        print(f"Added {delta_columns_added} delta columns including for dates not in current results")
        return delta_df
        
    except Exception as e:
        print(f"Error calculating deltas: {e}")
        import traceback
        traceback.print_exc()
        return None

################################################################################################
# Function 5: Save output files
def save_output(results_df, delta_df=None, result_csv_path=None):
    print("\n=== FUNCTION: SAVING OUTPUT ===")
    script_directory = os.path.dirname(__file__)
    
    if result_csv_path is None:
        result_csv_path = os.path.join(script_directory, "result_data.csv")
    
    # Save current results
    results_df.to_csv(result_csv_path, index=False)
    print(f"Results saved to {result_csv_path}")
    
    # Save delta comparison if available
    if delta_df is not None:
        delta_csv_path = os.path.join(script_directory, "result_data_with_delta.csv")
        delta_df.to_csv(delta_csv_path, index=False)
        print(f"Delta comparison saved to {delta_csv_path}")

################################################################################################
# Main function
print("\n=== STARTING SLOT CHECKING AUTOMATION ===")

# Load data
df_sql_data, previous_results, result_csv_path = load_data()

# Date information
now = datetime.datetime.now()
print(f"Current date: {calendar.month_name[now.month]} {now.day}, {now.year}")

# Initialize Chrome WebDriver
service = Service()
driver = webdriver.Chrome(service=service)

# Track success and failure counts
total_cases = len(df_sql_data)
successful_cases = 0
failed_cases = 0
failed_regions = []  # To track which regions failed

try:
    # Store results for all regions
    all_slot_data = []
    
    # Process each region in the dataframe
    for index, row in df_sql_data.iterrows():
        region = row['region']
        email = row['email']
        country = row['country']
        
        try:
            # Check slots for this region
            slot_data = check_slots_for_region(driver, region, email, country)
            all_slot_data.append(slot_data)
            
            # Check if we got meaningful results
            if slot_data and len(slot_data) > 0:
                successful_cases += 1
                print(f"✅ Successfully processed {region} with {len(slot_data)} days of slot data")
            else:
                failed_cases += 1
                failed_regions.append(region)
                print(f"❌ Failed to get slot data for {region}")
            
            # Add a small delay between requests
            time.sleep(1)
            
        except Exception as e:
            failed_cases += 1
            failed_regions.append(region)
            print(f"❌ Error processing {region}: {str(e)}")
            
            # Continue with next region even if this one fails
            continue
    
    # Process the results
    results_df, date_columns = process_results(all_slot_data, df_sql_data)
    
    # Calculate deltas if previous results exist
    delta_df = calculate_deltas(results_df, previous_results, date_columns)
    
    # Save output files
    save_output(results_df, delta_df, result_csv_path)
    
    # Generate and print summary report
    print("\n=== AUTOMATION SUMMARY ===")
    print(f"Total test cases: {total_cases}")
    print(f"Successfully processed: {successful_cases} ({successful_cases/total_cases*100:.1f}%)")
    print(f"Failed to process: {failed_cases} ({failed_cases/total_cases*100:.1f}%)")
    
    # Print details of failures if any
    if failed_cases > 0:
        print("\nFailed regions:")
        for region in failed_regions:
            print(f"  - {region}")
    
    print("\n=== AUTOMATION COMPLETED ===")
    
except Exception as e:
    print(f"\n❌ CRITICAL ERROR: {str(e)}")
finally:
    # Close the browser
    driver.quit()
