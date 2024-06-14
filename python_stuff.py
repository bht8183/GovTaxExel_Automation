from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import Workbook

# Specify the path to Edge WebDriver
service = Service('C:/Users/btokumoto/Downloads/edgedriver_win64/msedgedriver.exe')
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

        columns = row.find_all('td')
        col = [column.get_text().strip() for column in columns]
        #print("https://232app.azurewebsites.net/Forms/ExclusionRequestItem/"+col[0])
        if(col[0] != "No matching records found"):

            elements = []

            url2 = "https://232app.azurewebsites.net/Forms/ExclusionRequestItem/"+col[0]
            
            driver.switch_to.new_window('tab')
            driver.get(url2)

            html2 = driver.page_source # 2222222222222222222222222222222222222222222222222222222222
            soup2 = BeautifulSoup(html2, 'html.parser')

            time.sleep(3)

            Company_web_elem = driver.find_element(By.ID, 'BIS232Request_JSONData_RequestingOrg_WebsiteAddress')
            Company_web =  Company_web_elem.get_attribute('value')
            
            Prod_class_elem = driver.find_element(By.ID, 'BIS232Request_JSONData_MetalClass')
            Prod_class = Prod_class_elem.get_attribute('value')

            HTSUS_elem = driver.find_element(By.ID, 'BIS232Request_HTSUSCode')
            HTSUS = HTSUS_elem.get_attribute('value')

            Granted_ex_id_elem = driver.find_element(By.ID, 'BIS232Request_JSONData_PreviouslyGrantedER')
            Granted_ex_id = Granted_ex_id_elem.get_attribute('value')

            Requested_ex_elem = driver.find_element(By.ID, 'BIS232Request_JSONData_MetalClass')
            Requested_ex = Requested_ex_elem.get_attribute('value')

            Three_year_avr_elem = driver.find_element(By.ID, 'BIS232Request_JSONData_ExclusionExplanation_AvgAnnualConsumption')
            Three_year_avr = Three_year_avr_elem.get_attribute('value')

            Prod_desc_elem = driver.find_element(By.ID, 'BIS232Request_JSONData_ProductDescription_Description')
            Prod_desc = Prod_desc_elem.get_attribute('value')

            # missing commercial name

            Assoc_code_elem = driver.find_element(By.ID,'BIS232Request_JSONData_AdditionalDetails_AssociationCode')
            Assoc_code = Assoc_code_elem.get_attribute('value')

            Commercial_name_tab = driver.find_element(By.CLASS_NAME, 'tblcommercialnames')
            Commercial_names_row = Commercial_name_tab.find_elements(By.TAG_NAME,'tr')
            
            table3 = soup2.find('table', {"class":"tblcommercialnames"})
            print(table3.prettify)

            Commercial_names = ""
            
            counter = 0
            for elems in Commercial_names_row:
                #counter =+ 1
                try:
                    xval = driver.find_element(By.ID,"BIS232Request_JSONData_AdditionalDetails_CommercialNames_"+str(counter)+"_")
                except Exception as e:
                    print(e)
                    xval = driver.find_element(By.ID,"BIS232Request_JSONData_AdditionalDetails_CommercialNames_undefined_")
                Commercial_names += xval.get_attribute('value') + " "
                counter += 1
                
                #print(Commercial_names)
            
            Prod_applic_elem = driver.find_element(By.ID,'BIS232Request_JSONData_AdditionalDetails_ApplicationSuitability')
            Prod_applic = Prod_applic_elem.text

            #last_elements =  []
            #for row in Source_Countries_tab.find('tbody').find_all('tr'):
                #for elems in Commercial_names_row:
                #    last_elements.append(elems.get_attribute('value'))
                #print("aaaaaaaaaaaaaaaaaaaaaaaaaaa")



            


            elements.append(col[0])
            elements.append(col[6])
            elements.append(col[1])
            elements.append(Company_web)
            elements.append(HTSUS)
            elements.append(Prod_class)
            elements.append(Granted_ex_id)
            elements.append(Requested_ex)
            elements.append(Three_year_avr)
            elements.append(Prod_desc)
            elements.append(Commercial_names) # fix later
            elements.append(Assoc_code)
            elements.append(Prod_applic)
            #elements.append(last_elements)

            table2 = soup2.find('table', {"class":"table table-bordered bg-white tblsourcecountries"})
            #last_elements = []
            for raw in table2.find('tbody').find_all('tr'):
                
                print("yooooooooooooooooooooo")
                print(table2.prettify)
                ganggang = raw.find_all('td')
                #el = raw.find_element(By.ID,'BIS232Request_JSONData_SourceCountries_1__OriginCountry')
                #last_elements.append(el.text)
                #el = raw.find_element(By.ID,'BIS232Request_JSONData_SourceCountries_1__ExportCountry')
                #last_elements.append(el.text)
                #el = raw.find_element(By.ID,'BIS232Request_JSONData_SourceCountries_1__Manufacturer')
                #last_elements.append(el.text)
                #el = raw.find_element(By.ID,'BIS232Request_JSONData_SourceCountries_1__Supplier')
                #last_elements.append(el.text)


            #elements.append(last_elements)
            ws.append(elements)
            time.sleep(5)

            driver.close()
            driver.switch_to.window(original_window)

    product_filter1.clear()
    product_filter2.clear()

#for x in range(1,32):
#    read_page(x)
read_page(2)  

wb.save('tax_exemption_data.xlsx')
print('Data extraction and Excel writing complete.')