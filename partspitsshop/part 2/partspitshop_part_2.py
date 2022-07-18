from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import undetected_chromedriver as uc
import time
from openpyxl import Workbook


if __name__ == '__main__':
    options = Options()
    options.add_argument("--disable-extensions")
    options.binary_location = 'C:\Program Files\Google\Chrome Beta\Application\chrome.exe'
    driver = webdriver.Chrome(service=Service(ChromeDriverManager(version='104.0.5112.20').install()), options=options)
    driver.maximize_window()
    driver.implicitly_wait(30)

    final_list = []


    #First Website

    driver.get('https://www.partspitstoponline.com/sea-doo-oem-parts')

    years_containers = driver.find_elements(By.CSS_SELECTOR, "ul[class='partsubselect']")
    years_links = []
    years_values = []
    for year_container in years_containers:
        years_in_container = year_container.find_elements(By.TAG_NAME, 'a')
        for year in years_in_container:
            years_links.append(year.get_attribute('href'))
            years_values.append(year.text)

    for year_text, year_link in zip(years_values, years_links):
        driver.get(year_link)
        all_models = driver.find_elements(By.CSS_SELECTOR, "a[class='pjq']")
        for model in all_models:
            model_text = model.text
            list_to_append = ["Jet Ski", "Sea Doo", year_text, model_text]
            print(list_to_append)
            final_list.append(list_to_append)

    #Second website

    driver.get('https://www.kawasakipartspitstop.com/kawasaki-jet-ski-parts')

    years_containers = driver.find_elements(By.CSS_SELECTOR, "ul[class='partsubselect']")
    years_links = []
    years_values = []
    for year_container in years_containers:
        years_in_container = year_container.find_elements(By.TAG_NAME, 'a')
        for year in years_in_container:
            years_links.append(year.get_attribute('href'))
            years_values.append(year.text)

    for year_text, year_link in zip(years_values, years_links):
        driver.get(year_link)
        all_models = driver.find_elements(By.CSS_SELECTOR, "a[class='pjq']")
        for model in all_models:
            model_text = model.text
            list_to_append = ["Jet Ski", "Kawasaki", year_text, model_text]
            print(list_to_append)
            final_list.append(list_to_append)


    #Third website

    types_links = []
    types_names = []

    driver.get('https://www.partspitstop.com/oem-polaris-parts')

    types_container = driver.find_element(By.CSS_SELECTOR,"div[id='container_5724']")

    types_wrappers = types_container.find_elements(By.CSS_SELECTOR,"div[class='grid_8']")[:4]

    for type_wrapper in types_wrappers:
        types_links.append(type_wrapper.find_element(By.CSS_SELECTOR,"a[class='button button_16']").get_attribute('href'))
        if "atv" in type_wrapper.find_element(By.CSS_SELECTOR,"a[class='button button_16']").text.lower():
            types_names.append("ATV")
        else:
            types_names.append("UTV")

    for type_link,types_name in zip(types_links,types_names):

        driver.get(type_link)

        years_containers = driver.find_elements(By.CSS_SELECTOR, "ul[class='partsubselect']")
        years_links = []
        years_values = []
        for year_container in years_containers:
            years_in_container = year_container.find_elements(By.TAG_NAME, 'a')
            for year in years_in_container:
                years_links.append(year.get_attribute('href'))
                years_values.append(year.text)

        for year_text, year_link in zip(years_values, years_links):
            if "serie" in year_text.lower():
                pass
            else:
                driver.get(year_link)
                all_models = driver.find_elements(By.CSS_SELECTOR, "a[class='pjq']")
                for model in all_models:
                    model_text = model.text
                    list_to_append = [types_name, "Polaris", year_text, model_text]
                    print(list_to_append)
                    final_list.append(list_to_append)

    #Fourth Website

    driver.get('https://www.partspitstop.com/yamaha-waverunner-parts')

    years_containers = driver.find_elements(By.CSS_SELECTOR, "ul[class='partsubselect']")
    years_links = []
    years_values = []
    for year_container in years_containers:
        years_in_container = year_container.find_elements(By.TAG_NAME, 'a')
        for year in years_in_container:
            years_links.append(year.get_attribute('href'))
            years_values.append(year.text)

    for year_text, year_link in zip(years_values, years_links):
        driver.get(year_link)
        all_models = driver.find_elements(By.CSS_SELECTOR, "a[class='pjq']")
        for model in all_models:
            model_text = model.text
            list_to_append = ["Jet Ski", "Yamaha", year_text, model_text]
            print(list_to_append)
            final_list.append(list_to_append)

    workbook_name = 'Final_results_of_scraping.xlsx'
    wb = Workbook()
    page = wb.active
    headers = ['Type', 'Make', 'Year', 'Model']
    page.append(headers)
    for row in final_list:
        page.append(row)
    wb.save(filename='Final_results_of_scraping.xlsx')







