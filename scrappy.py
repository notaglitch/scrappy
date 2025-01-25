import requests
from bs4 import BeautifulSoup
import pandas as pd

url = 'https://www.orientation.ch/dyn/show/170785?lang=fr&Idx=0&OrderBy=1&Order=0&PostBackOrder=0&postBack=true&CountResult=0&Total_Idx=0&CounterSearch=4&UrlAjaxWebSearch=%2FLeFiWeb%2FAjaxWebSearch&getTotal=False&isBlankState=True&prof_=88613.1&fakelocalityremember=&LocName=Neuch%C3%A2tel&LocId=neuchatel-ne-ch&Area=10&AreaCriteria=10'

response = requests.get(url)

soup = BeautifulSoup(response.text, 'html.parser')

job_listings = soup.find_all('div', class_='display-table-row')

companies = []
job_titles = []
locations = []
languages = []

for job in job_listings:
    if 'head-part' in job.get('class', []):
        continue
    if 'result-info' not in job.get('class', []):
        company = job.find('div', class_='table-col-1')
        if company:
            companies.append(company.text.strip())
        else:
            companies.append("N/A")

        job_title = job.find('div', class_='table-col-2')
        if job_title:
            job_titles.append(job_title.text.strip())
        else:
            job_titles.append("N/A")

        location = job.find('div', class_='table-col-3')
        if location:
            locations.append(location.text.strip())
        else:
            locations.append("N/A")

        language = job.find('div', class_='table-col-4')
        if language:
            languages.append(language.text.strip())
        else:
            languages.append("N/A")

data = {
    companies[0]: companies,
    job_titles[0]: job_titles,
    locations[0]: locations,
    languages[0]: languages
}

df = pd.DataFrame(data)

df.to_csv('job_listing.csv', index=False)

print("Data has been saved successfully")
