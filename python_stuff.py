<<<<<<< HEAD
from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import Workbook

# Specify the path to Edge WebDriver
service = Service('C:/Users/hiron/edgedriver_win64/msedgedriver.exe')
driver = webdriver.Edge(service=service)

url = 'https://232app.azurewebsites.net/'

# Open the URL using Edge
driver.get(url)
original_window = driver.current_window_handle

time.sleep(5)  # Sleep for 5 seconds;


data = []
wb = Workbook()
ws = wb.active
ws.title = "Tax Exemption Data"
    
# Save the workbook to an Excel file
#wb.save('tax_exemption_data.xlsx')
print('Data extraction and Excel writing complete.')

def read_page(x):
    # Find the product filter input and enter 'Aluminum'
    product_filter1 = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Product']")
    product_filter1.send_keys("Aluminum")
    product_filter1.send_keys(Keys.RETURN)

    product_filter2 = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Posted Date']")
    str1 = "5/" + str(x) + "/2024"
    product_filter2.send_keys(str1)
    product_filter2.send_keys(Keys.RETURN)

    
    time.sleep(5)

    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table', id='erdashboard')

    for row in table.find('tbody').find_all('tr'):
        #git print (row)
        columns = row.find_all('td')
        col = [column.get_text().strip() for column in columns]
        print("https://232app.azurewebsites.net/Forms/ExclusionRequestItem/"+col[0])
        if(col[0] != "No matching records found"):

            driver.switch_to.new_window('tab')
            driver.get("https://232app.azurewebsites.net/Forms/ExclusionRequestItem/"+col[0])
            ws.append(col)
            time.sleep(5)
            driver.close()
            driver.switch_to.window(original_window)
            #driver.close()

    product_filter1.clear()
    product_filter2.clear()

for x in range(1,4):
    read_page(x)

#read_page(1)   

wb.save('tax_exemption_data.xlsx')
print('Data extraction and Excel writing complete.')