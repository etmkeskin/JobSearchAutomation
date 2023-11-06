from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl

# Set the path to the Firefox executable
firefox_path = "C:\\Program Files\\Mozilla Firefox\\firefox.exe"

# Set the URL of the job search website
url = "https://www.workopolis.com/en/"

# Set the path to the Excel file where you want to save the job data
excel_file_path = "C:\\Users\\etmke\\Downloads\\jobSearch.xlsx"

# Set the desired job search parameters
location = "Toronto"
job_titles = ["software developer", "business analyst", "project management", "mobile developer"]

# Initialize the Firefox WebDriver with the specified executable path
firefox_options = webdriver.FirefoxOptions()
firefox_options.binary_location = firefox_path
driver = webdriver.Firefox(options=firefox_options)

# Open the website
driver.get(url)

# Locate the "location" input field using its id attribute
location_input = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, 'location-input'))
)

location_input.clear()
location_input.send_keys(location)
location_input.send_keys(Keys.RETURN)

# Wait for the search results to load
time.sleep(5)

# Create a new Excel workbook
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.append(["Job Title", "Company Name", "Job Description", "Salary"])

# Loop through the job titles and scrape job data
for job_title in job_titles:
    # Wait for the "query-input" input field to be present
    job_search_input = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "query-input"))
    )

    job_search_input.clear()
    job_search_input.send_keys(job_title)
    job_search_input.send_keys(Keys.RETURN)

    # Wait for the search results to load
    time.sleep(5)

    # Locate the job listings container using the XPath
    job_list_xpath = '//*[@id="job-list"]'
    job_list_element = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, job_list_xpath))
    )

    # Scrape job data and add it to the Excel worksheet
    job_elements = job_list_element.find_elements(By.CLASS_NAME, "job-card")
    for job_element in job_elements:
        job_title = job_element.find_element(By.CLASS_NAME, "job-title").text
        company_name = job_element.find_element(By.CLASS_NAME, "job-company").text
        job_description = job_element.find_element(By.CLASS_NAME, "job-description").text
        salary = job_element.find_element(By.CLASS_NAME, "job-salary").text if job_element.find_elements(By.CLASS_NAME,
                                                                                                         "job-salary") else "N/A"
        print(f"Job Title: {job_title}")
        print(f"Company Name: {company_name}")
        print(f"Job Description: {job_description}")
        print(f"Salary: {salary}")
        worksheet.append([job_title, company_name, job_description, salary])

# Save the Excel file
workbook.save(excel_file_path)

# Close the browser
driver.quit()