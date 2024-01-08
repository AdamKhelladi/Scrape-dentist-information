# Project One : 

from selenium import webdriver 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC

import requests
import time
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd 

options = Options()
options.add_experimental_option("detach", True)

path = r"C:\Users\hp\AppData\Local\Programs\Python\Python311\python web scraping\selenium\chromedriver-win64\chromedriver.exe"

service = Service(path)

driver = webdriver.Chrome(path, options=options, service=service)
driver.get("https://www.google.com/maps/@36.7099904,5.0495488,13z?entry=ttu")

driver.maximize_window()
time.sleep(3)

search = driver.find_element(By.NAME, "q")
search.clear()
search.send_keys("dentist new york")
search.send_keys(Keys.RETURN)

wait = WebDriverWait(driver, 10)
dentists = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "Nv2PK")))

names = []
locations = []
ranks = []
websites = []

for dentist in dentists:
  try:
    dentist.click()

    dentist_name = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "lfPIob")))
    print(f"- Dentist Name: {dentist_name.text}")
    names.append(dentist_name.text)
    time.sleep(3)

    location_entry = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "RcCsl")))
    dentist_location = location_entry.find_element(By.CLASS_NAME, "rogA2c")
    print(f"- Dentist Location: {dentist_location.text}")
    locations.append(dentist_location.text)
    time.sleep(3)

    dentist_rank = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "F7nice")))
    print(f"- Dentist Rank: {dentist_rank.text}")
    ranks.append(dentist_rank.text)
    time.sleep(3)

    dentist_website = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ITvuef ")))
    print(f"- Dentist WebSite: {dentist_website.text}")
    websites.append(dentist_website.text)
    time.sleep(3)

    print("-" * 20)
    print()
    time.sleep(5)

  except Exception:
    continue

  finally:
    driver.back()

excel_file_path= r"C:\Users\hp\AppData\Local\Programs\Python\Python311\python web scraping\Projects\dentists_info.xlsx"

# Create a DataFrame from the lists
data = {'Dentist Name': names, 'Location': locations, 'Rank': ranks, 'Website': websites}
df = pd.DataFrame(data)

# Save the DataFrame to an Excel file
df.to_excel(excel_file_path, index=False)
print(f'Data has been saved to {excel_file_path}')