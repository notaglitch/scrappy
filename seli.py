from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook

chromedriver_path = r"C:\Program Files\Google\Chrome\Driver\chromedriver-win64\chromedriver.exe"
service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service)

url = "https://www.orientation.ch/dyn/show/170785?lang=fr&Idx=20&OrderBy=1&Order=0&PostBackOrder=0&postBack=true&CountResult=446&Total_Idx=446&CounterSearch=26&UrlAjaxWebSearch=%2FLeFiWeb%2FAjaxWebSearch&getTotal=True&isBlankState=False&prof_=68800.1&fakelocalityremember=&LocName=Neuch%C3%A2tel&LocId=neuchatel-ne-ch&Area=10&AreaCriteria=10&langcode_=de&langcode_=fr&langcode_=it&langcode_=rm&langcode_=en"
driver.get(url)

WebDriverWait(driver, 10).until(EC.presence_of_element_located(
    (By.XPATH, "//a[@id='aSearchPaging']")))


def decrypt_email(encrypted_str, encryption_key="EncryptionKey"):
    encrypted_numbers = encrypted_str.split(",")
    decrypted_str = ""
    t = 0

    for num in encrypted_numbers:
        u = ord(encryption_key[t % len(encryption_key)]) % 96
        r = int(num) - u

        if r < 32:
            r += 96

        decrypted_str += chr(r)
        t += 1

    return decrypted_str


def scroll_to_bottom():
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        # Leave 1500px at the bottom
        new_scroll_position = last_height - 1500
        driver.execute_script(f"window.scrollTo(0, {new_scroll_position});")
        time.sleep(3)

        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height


def click_more_results():
    while True:
        try:
            scroll_to_bottom()
            more_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[@id='aSearchPaging' and contains(text(), 'Afficher plus de résultats')]"))
            )

            more_button.click()
            print("Clicked 'More' button...")

            time.sleep(5)
        except Exception as e:
            print("No more 'More' buttons to click.")
            break


click_more_results()

page_source = driver.page_source

soup = BeautifulSoup(page_source, 'html.parser')

job_listings = soup.find_all('div', class_='display-table-row')

companies = []
job_titles = []
locations = []
languages = []
addresses = []
emails = []

for job in job_listings:
    if 'result-info' not in job.get('class', []):
        company = job.find('div', class_='table-col-1')
        if company:
            if company.text != "Entreprise":
                companies.append(company.text.strip())
        else:
            companies.append("N/A")

        job_title = job.find('div', class_='table-col-2')
        if job_title:
            if job_title.text != "Profession":
                job_titles.append(job_title.text.strip())
        else:
            job_titles.append("N/A")

        location = job.find('div', class_='table-col-3')
        if location:
            if location.text != "Localité":
                locations.append(location.text.strip())
        else:
            locations.append("N/A")

        language = job.find('div', class_='table-col-4')
        if language:
            if language.text != "Langue":
                languages.append(language.text.strip())
        else:
            languages.append("N/A")

job_contacts = soup.find_all('div', class_='result-elem-lower')
for result in job_contacts:
    email_scripts = soup.find_all(
        'script', string=lambda s: s and 'sdbb.DecrypteEmail' in s)
    if not email_scripts:
        email_scripts = soup.find_all(
            'script', text=lambda t: t and 'sdbb.DecryptEmail' in t
        )

    address_section = result.find('div', class_='w45')
    if address_section:
        address = address_section.find('p')
        if address:
            address_text = address.get_text(separator=" ").strip()
            addresses.append(address_text)
        else:
            addresses.append("N/A")
    else:
        addresses.append("N/A")

emails = []

for script in email_scripts:
    if script:
        script_text = script.string.strip()
        last_numbers = script_text.split("'")[5]
        decrypted_email = decrypt_email(last_numbers)
        emails.append(decrypted_email)
    else:
        emails.append("NOT FOUND")

data = {
    "Company": companies,
    "Title": job_titles,
    "Location": locations,
    "Language": languages,
}

df = pd.DataFrame(data)
df.to_excel('job_listing.xlsx', index=False, header=True, engine='openpyxl')

wb = load_workbook('job_listing.xlsx')
ws = wb.active

for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column].width = adjusted_width

file_name = input("Give a name for this file: ")
wb.save(f"{file_name}.xlsx")

print("Data has been saved successfully")
driver.quit()
