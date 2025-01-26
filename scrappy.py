import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook


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


url = 'https://www.orientation.ch/dyn/show/170785?lang=fr&Idx=20&OrderBy=1&Order=0&PostBackOrder=0&postBack=true&CountResult=446&Total_Idx=446&CounterSearch=8&UrlAjaxWebSearch=%2FLeFiWeb%2FAjaxWebSearch&getTotal=True&isBlankState=False&prof_=68800.1&fakelocalityremember=&LocName=Neuch%C3%A2tel&LocId=neuchatel-ne-ch&Area=10&LocatorVM.Area=10&langcode_=de&langcode_=fr&langcode_=it&langcode_=rm&langcode_=en'

response = requests.get(url)

soup = BeautifulSoup(response.text, 'html.parser')

job_listings = soup.find_all('div', class_='display-table-row')

companies = []
job_titles = []
locations = []
languages = []
addresses = []

for job in job_listings[1:]:
    if 'result-info' not in job.get('class', []):
        company = job.find(
            'div', class_='table-col-1')
        companies.append(company.text.strip() if company else "N/A")

        job_title = job.find(
            'div', class_='table-col-2')
        job_titles.append(job_title.text.strip() if job_title else "N/A")

        location = job.find(
            'div', class_='table-col-3')
        locations.append(location.text.strip() if location else "N/A")

        language = job.find(
            'div', class_='table-col-4')
        languages.append(language.text.strip() if language else "N/A")

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


keys = []

for script in email_scripts:
    if script:
        script_text = script.string.strip()
        last_numbers = script_text.split("'")[5]
        keys.append(last_numbers)
    else:
        keys.append("NOT FOUND")

for i, encrypted in enumerate(keys, 1):
    decrypted_email = decrypt_email(encrypted)
    print(f"Key {i}: {decrypted_email}")

data = {
    "Company": companies,
    "Title": job_titles,
    "Location": locations,
    "Language": languages,
    "Address": addresses
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

wb.save('job_listing.xlsx')

print("Data has been saved successfully")
