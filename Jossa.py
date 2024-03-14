import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Start a WebDriver session with Edge
driver = webdriver.Edge()
driver.get('https://josaa.admissions.nic.in/applicant/SeatAllotmentResult/CurrentORCR.aspx')

try:
    # Define reusable wait object with a timeout of 10 seconds
    wait = WebDriverWait(driver, 10)

    # Selecting round 6
    round_option = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ddlroundno_chosen"]/div/ul/li[7]')))
    round_option.click()

    # Selecting Institute Type
    institute_type_option = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ddlInstype_chosen"]/div/ul/li[4]')))
    institute_type_option.click()

    # Selecting Institute Name
    Ins_name_opt = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ddlInstitute_chosen"]/div/ul/li[2]')))
    Ins_name_opt.click()

    # Selecting Program
    program_opt = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ddlBranch_chosen"]/div/ul/li[2]')))
    program_opt.click()

    # Selecting Category
    category_opt = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ddlSeattype_chosen"]/div/ul/li[2]')))
    category_opt.click()

    # Clicking Submit button
    Submit = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_btnSubmit"]')))
    Submit.click()

    # Wait for the table data to load
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_GridView1 > tbody > tr')))

    # Extract data from the table
    data = []
    Institutes = driver.find_elements(By.CSS_SELECTOR, '#ctl00_ContentPlaceHolder1_GridView1 > tbody > tr')
    for row in Institutes:
        cols = row.find_elements(By.TAG_NAME, 'td')
        row_data = [col.text for col in cols]
        data.append(row_data)

    # Create a DataFrame
    headers = ["Institute", "Branch", "Quota", "Category", "Gender", "Open Rank", "Close Rank"]
    First_allocation = pd.DataFrame(data, columns=headers)

    # Save the data to CSV
    First_allocation.to_csv('Josaa_OR_CR_Round6.csv', index=False)

except TimeoutException as e:
    print("Timeout occurred while waiting for elements to load:", e)

except Exception as e:
    print("An error occurred:", e)

finally:
    # Closing the webdriver
    driver.quit()