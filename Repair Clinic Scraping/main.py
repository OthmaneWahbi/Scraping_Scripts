from selenium.webdriver.common.by import By
import undetected_chromedriver as uc
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from openpyxl import Workbook
from openpyxl import load_workbook

if __name__=='__main__':
    options = Options()
    options.add_argument("--disable-extensions")
    options.binary_location = 'C:\Program Files\Google\Chrome Beta\Application\chrome.exe'
    driver = uc.Chrome(service=Service(ChromeDriverManager(version='104.0.5112.20').install()), options=options)
    driver.implicitly_wait(30)
    driver.get('https://www.repairclinic.com/')


    workbook_name = 'Final_results_of_scraping.xlsx'
    wb = Workbook()
    page = wb.active
    headers = ['Type', 'Make', 'Model']
    page.append(headers)


    appliances_Container = driver.find_element(By.CSS_SELECTOR,"div[class='jsx-1325188385 applicancesContainer']")
    appliances = appliances_Container.find_elements(By.CSS_SELECTOR, "div[class='jsx-3520301984 applianceLink']")

    final_final_list=[]

    links_type = []
    types = []
    for appliance in appliances:
        if appliance.find_element(By.CSS_SELECTOR,"div[class='jsx-3520301984 lbl']").text == "Small Engine" or appliance.find_element(By.CSS_SELECTOR,"div[class='jsx-3520301984 lbl']").text == "Microwave":
            pass
        else:
            links_type.append(appliance.find_element(By.TAG_NAME,"a").get_attribute('href'))
            types.append(appliance.find_element(By.CSS_SELECTOR,"div[class='jsx-3520301984 lbl']").text)

    for link_type,type_text in zip(links_type,types):
        driver.get(link_type)
        print(type_text)

        brand_names_container = driver.find_element(By.CSS_SELECTOR, "div[id='linklist_facet_brand_names']")
        brands = brand_names_container.find_elements(By.TAG_NAME, 'a')
        makes = []
        links_make = []
        for brand in brands:
            links_make.append(brand.get_attribute('href'))
            makes.append(brand.text)
        for link_make, make_text in zip(links_make, makes):
            driver.get(link_make)
            print(make_text)
            try:
                popular_models = driver.find_element(By.CSS_SELECTOR, "div[id='sectionPanel_popularModels']")
                final_list = []
                try:
                    driver.get(popular_models.find_element(By.XPATH, "//*[contains(text(), 'View all models')]").get_attribute('href'))
                    test = True
                    while test == True:
                        models_section = driver.find_element(By.CSS_SELECTOR, "div[class='modelsSection']")
                        all_models_inpage = driver.find_elements(By.CSS_SELECTOR, "div[class='modelLink col-md-4 col-sm-6 col-xs-12']")
                        for model in all_models_inpage:
                            try:
                                model = model.find_element(By.TAG_NAME, "a")
                            except:
                                pass
                        for model in all_models_inpage:
                            try:
                                model_text = model.text.split()
                            except:
                                pass
                            try:
                                model_text.remove(type_text)
                            except:
                                words_to_remove = type_text.split()
                                for word in words_to_remove:
                                    try:
                                        model_text.remove(word)
                                    except:
                                        pass
                            try:
                                model_text.remove(make_text)
                            except:
                                words_to_remove = make_text.split()
                                for word in words_to_remove:
                                    try:
                                        model_text.remove(word)
                                    except:
                                        pass
                            model_final_text = ''
                            for x in model_text:
                                model_final_text = model_final_text + ' ' +  x
                            if "Range/Stove/Oven" in model_final_text:
                                model_final_text = model_final_text.replace('Range/Stove/Oven', '')
                            list_to_append = [type_text, make_text, model_final_text.lstrip()]
                            if list_to_append in final_list:
                                pass
                            else:
                                final_list.append(list_to_append)
                                print(list_to_append)
                                page.append(list_to_append)
                        try:
                            next_button = driver.find_element(By.CSS_SELECTOR,"li[class='next']")
                            next_button = next_button.find_element(By.TAG_NAME,'a')
                            next_button.click()
                            #driver.refresh()
                            time.sleep(1)
                        except:
                            next_button = driver.find_element(By.CSS_SELECTOR, "li[class='next disabled']")
                            test = False
                except:
                    list_of_models = popular_models.find_element(By.CSS_SELECTOR,"div[class='linkList']")
                    models = list_of_models.find_elements(By.TAG_NAME,"a")
                    for model in models:
                        model_text = model.text
                        list_to_append = [type_text,make_text,model_text]
                        if list_to_append in final_list:
                            pass
                        else:
                            final_list.append(list_to_append)
                            print(list_to_append)
                            page.append(list_to_append)
                #final_final_list = final_final_list + final_list
                wb.save(filename='Final_results_of_scraping.xlsx')
            except:
                pass


