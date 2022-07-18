from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import undetected_chromedriver as uc
#from selenium.webdriver.support.ui import Select

from selenium.webdriver.support.select import Select

import time


from openpyxl import Workbook
from openpyxl import load_workbook

if __name__=='__main__':
    uc.TARGET_VERSION = 104
    driver = uc.Chrome()
    driver.implicitly_wait(10)
    driver.get('https://www.hlsproparts.com/chainsaw-parts-s/42.htm')

    brand_button = driver.find_element(By.CSS_SELECTOR, "span[id='brand-selector-button']")
    product_button = driver.find_element(By.CSS_SELECTOR, "span[id='product-selector-button']")
    model_button = driver.find_element(By.CSS_SELECTOR, "span[id='model-selector-button']")

    time.sleep(2)
    brand_button.click()
    brand_selector_menu = driver.find_element(By.CSS_SELECTOR, "ul[id='brand-selector-menu']")
    brands = brand_selector_menu.find_elements(By.CSS_SELECTOR,"li[class='ui-menu-item']")
    time.sleep(2)
    brand_button.click()
    final_list = []

    for brand in brands:
        brand_button.send_keys(Keys.ARROW_DOWN)
        brand_text = brand_button.text

        time.sleep(2)
        product_button.click()
        product_selector_menu = driver.find_element(By.CSS_SELECTOR, "ul[id='product-selector-menu']")
        products = product_selector_menu.find_elements(By.CSS_SELECTOR, "li[class='ui-menu-item']")
        time.sleep(2)
        product_button.click()

        for product in products:
            product_button.send_keys(Keys.ARROW_DOWN)
            product_text = product_button.text
            time.sleep(2)
            model_button.click()
            model_selector_menu = driver.find_element(By.CSS_SELECTOR, "ul[id='model-selector-menu']")
            models = model_selector_menu.find_elements(By.CSS_SELECTOR, "li[class='ui-menu-item']")
            time.sleep(2)
            model_button.click()
            for model in models:
                model_button.send_keys(Keys.ARROW_DOWN)
                model_text = model_button.text
                if model_text == "SHOW ALL":
                    pass
                else:
                    list_to_append = [brand_text,product_text,model_text]
                    print(list_to_append)
                    final_list.append(list_to_append)

    workbook_name = 'Final_results_of_scraping.xlsx'
    wb = Workbook()
    page = wb.active
    headers = ['Type', 'Make', 'Model']
    page.append(headers)
    for row in final_list:
        page.append(row)
    wb.save(filename='Final_results_of_scraping.xlsx')





