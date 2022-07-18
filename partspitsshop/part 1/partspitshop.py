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
    driver.get('https://www.partspitstop.com/snowmobile-oem-parts')

    final_list = []

    makes_container = driver.find_element(By.CSS_SELECTOR,"div[id='container_26569']")
    makes = makes_container.find_elements(By.CSS_SELECTOR,"div[class='contentwrapper']")

    make_links = []
    make_names = []

    for make in makes:
        a_tag_element = make.find_elements(By.TAG_NAME,"p")[1].find_element(By.TAG_NAME,"a")
        make_links.append(a_tag_element.get_attribute('href'))
        if "polaris" in a_tag_element.text.lower():
            make_names.append("Polaris")
        elif "ski" in a_tag_element.text.lower():
            make_names.append("Ski-doo")
        elif "yamaha" in a_tag_element.text.lower():
            make_names.append("Yamaha")

    for make_link, make_name in zip(make_links,make_names):
        driver.get(make_link)
        years_containers = driver.find_elements(By.CSS_SELECTOR, "ul[class='partsubselect']")
        years_links = []
        years_values = []
        for year_container in years_containers:
            years_in_container = year_container.find_elements(By.TAG_NAME,'a')
            for year in years_in_container:
                years_links.append(year.get_attribute('href'))
                years_values.append(year.text)

        for year_text,year_link in zip(years_values,years_links):
            driver.get(year_link)
            all_models = driver.find_elements(By.CSS_SELECTOR,"a[class='pjq']")
            for model in all_models:
                model_text = model.text
                list_to_append = ["Snowmobile",make_name,year_text,model_text]
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













