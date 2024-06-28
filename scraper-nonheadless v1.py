import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# Function to initialize the WebDriver
def init_driver():
    options = Options()
    # options.add_argument('--headless')  # Uncomment this line for headless operation
    options.add_argument('--disable-gpu')
    service = Service('C:\\chromedriver-win64\\chromedriver.exe')  # Updated Chrome driver path
    driver = webdriver.Chrome(service=service, options=options)
    return driver

# Function to log in to the website
def login(driver, username, password):
    driver.get("https://student.mitapps.in/landing")  # Navigate to landing page
    time.sleep(2)
    
    try:
        # Click on Student Login
        student_login_button = driver.find_element(By.XPATH, "//a[contains(text(), 'Student Login')]")
        student_login_button.click()
        
        print(f"Trying Username - '{username}' and Password - '{password}'")
        
        username_input = driver.find_element(By.XPATH, "//input[@placeholder='Username']")
        password_input = driver.find_element(By.XPATH, "//input[@placeholder='Password']")
        
        username_input.send_keys(username)
        password_input.send_keys(password)
        password_input.send_keys(Keys.RETURN)
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(@href, 'student_result')]"))
        )
        return True
    except NoSuchElementException as e:
        print(f"Error in login: {e}")
        return False
    except TimeoutException:
        print("Timeout while logging in.")
        return False

# Function to scrape the SGPA and CGPA
def scrape_results(driver):
    driver.get("https://student.mitapps.in/student_result")
    time.sleep(2)
    
    try:
        sgpa_elements = driver.find_elements(By.XPATH, "//div[contains(text(), 'SGPA')]/following-sibling::div")
        sgpa = [el.text for el in sgpa_elements]
        
        cgpa_element = driver.find_element(By.XPATH, "//div[contains(text(), 'CGPA')]/following-sibling::div")
        cgpa = cgpa_element.text
        
        return sgpa, cgpa
    except NoSuchElementException as e:
        print(f"Error in scraping results: {e}")
        return [], None

# Main function to read input, process each user, and save results
def main(input_excel, output_excel):
    driver = init_driver()
    
    input_data = pd.read_excel(input_excel, header=None)  # Read without assuming headers
    
    results = []

    for index, row in input_data.iterrows():
        username = row.iloc[0]  # Username from first column
        password = row.iloc[1]  # Password from second column
        
        print(f"Processing Username - '{username}' and Password - '{password}'")
        
        if login(driver, username, password):
            sgpa, cgpa = scrape_results(driver)
            results.append({
                "Username": username,
                "Password": password,
                "SGPA": sgpa,
                "CGPA": cgpa
            })
        else:
            results.append({
                "Username": username,
                "Password": password,
                "SGPA": None,
                "CGPA": None
            })
    
    driver.quit()
    
    output_data = pd.DataFrame(results)
    output_data.to_excel(output_excel, index=False)

# Run the main function
if __name__ == "__main__":
    input_excel = "D:\\MIT Result Scraper\\Blank.xlsx"  # Updated input Excel path
    output_excel = "D:\\MIT Result Scraper\\cute.xlsx"  # Updated output Excel path
    main(input_excel, output_excel)
