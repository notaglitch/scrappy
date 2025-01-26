from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from bs4 import BeautifulSoup

# Set the path to the chromedriver
chromedriver_path = r"C:\Program Files\Google\Chrome\Driver\chromedriver-win64\chromedriver.exe"
service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service)

# Open the website
url = 'https://www.orientation.ch/dyn/show/170785?lang=fr&Idx=20&OrderBy=1&Order=0&PostBackOrder=0&postBack=true&CountResult=27&Total_Idx=27&CounterSearch=10&UrlAjaxWebSearch=%2FLeFiWeb%2FAjaxWebSearch&getTotal=True&isBlankState=False&prof_=47419.1&fakelocalityremember=&LocName=Neuch%C3%A2tel&LocId=neuchatel-ne-ch&Area=10&LocatorVM.Area=10&langcode_=de&langcode_=fr&langcode_=it&langcode_=rm&langcode_=en'
driver.get(url)

# Wait for the page to load
WebDriverWait(driver, 10).until(EC.presence_of_element_located(
    (By.XPATH, "//a[@id='aSearchPaging']")))

# Function to click the "More results" button until it disappears


def click_more_results():
    while True:
        try:
            # Wait for the "More results" button to be clickable
            more_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[@id='aSearchPaging' and contains(text(), 'Afficher plus de résultats')]"))
            )

            # Click the "More" button
            more_button.click()
            print("Clicked 'More' button...")

            # Wait for new results to load
            time.sleep(3)
        except Exception as e:
            # Break the loop if the "More" button is not found (i.e., no more results to load)
            print("No more 'More' buttons to click.")
            break


# Call the function to click the "More results" button
click_more_results()

# Once all content is loaded, get the page source
page_source = driver.page_source

# Parse the page with BeautifulSoup
soup = BeautifulSoup(page_source, 'html.parser')

# Now scrape the data from the page
job_listings = soup.find_all('div', class_='display-table-row')

companies = []
job_titles = []
locations = []
languages = []

for job in job_listings[1:]:
    if 'result-info' not in job.get('class', []):
        company = job.find('div', class_='table-col-1')
        companies.append(company.text.strip() if company else "N/A")

        job_title = job.find('div', class_='table-col-2')
        job_titles.append(job_title.text.strip() if job_title else "N/A")

        location = job.find('div', class_='table-col-3')
        locations.append(location.text.strip() if location else "N/A")

        language = job.find('div', class_='table-col-4')
        languages.append(language.text.strip() if language else "N/A")

email_scripts = soup.find_all(
    'script', string=lambda s: s and 'sdbb.DecrypteEmail' in s)
if not email_scripts:
    email_scripts = soup.find_all(
        'script', text=lambda t: t and 'sdbb.DecryptEmail' in t
    )

keys = []

for script in email_scripts:
    if script:
        script_text = script.string.strip()
        last_numbers = script_text.split("'")[5]
        keys.append(last_numbers)
    else:
        keys.append("NOT FOUND")

for i, encrypted in enumerate(keys, 1):
    print(f"Key {i}: {encrypted}")
# Print the extracted data
for i in range(len(companies)):
    print(
        f"Company: {companies[i]}, Title: {job_titles[i]}, Location: {locations[i]}, Language: {languages[i]}")

# Close the browser
driver.quit()
