import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

import os
import sys
from configparser import ConfigParser

from openpyxl import Workbook


def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS

    except Exception:
        base_path = os.path.dirname(__file__)

    return os.path.join(base_path, relative_path)

class Equipement(webdriver.Chrome):
    def __init__(self, driver_path = r"./driver", teardown = False):
        config = ConfigParser()
        config.read("config.ini")
        options = Options()
        options.binary_location = config.get("chrome", "chrome_location")
        options.add_experimental_option('excludeSwitches',['enable-logging'])
        self.driver_path = resource_path(config.get("chromedriver", "path"))
        self.teardown = teardown
        os.environ['PATH'] += self.driver_path
        super(Equipement, self).__init__(chrome_options=options)
        self.implicitly_wait(60)
        self.maximize_window()

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.teardown:
            self.quit()

    def land_first_page(self):
        config = ConfigParser()
        config.read("config.ini")
        self.get(config.get('website','url'))


    def get_data_treadmill(self,finallist=[],url='https://www.fitnessrepairparts.com/type/View/1/Treadmills'):
        self.get(url)
        typetext = "Treadmill"
        time.sleep(2)
        main_content = self.find_element(By.TAG_NAME,"tbody")
        table_rows = main_content.find_elements(By.TAG_NAME, "tr")
        for row in table_rows:
            all_table_data = row.find_elements(By.TAG_NAME, "td")
            for td in all_table_data:
                try :
                    firsta = td.find_element(By.TAG_NAME,"a")
                    all_models = td.find_elements(By.TAG_NAME, "a")
                    if "Show More" in firsta.text:
                        firsta.click()
                        all_models = all_models[1:len(all_models) - 1]
                    maketext = td.find_element(By.TAG_NAME,'b').text
                    maketext = maketext.split('\n')[0]
                    for model in all_models:
                        modeltext = model.text
                        listtoappend = [typetext,maketext,modeltext]
                        print(listtoappend)
                        finallist.append(listtoappend)
                except:
                    break
        return finallist

    def get_data_elipticals(self,finallist=[],url='https://www.fitnessrepairparts.com/type/View/2/Ellipticals'):
        self.get(url)
        typetext = "Elipticals"
        time.sleep(2)
        main_content = self.find_element(By.TAG_NAME,"tbody")
        table_rows = main_content.find_elements(By.TAG_NAME, "tr")
        for row in table_rows:
            all_table_data = row.find_elements(By.TAG_NAME, "td")
            for td in all_table_data:
                try :
                    firsta = td.find_element(By.TAG_NAME,"a")
                    all_models = td.find_elements(By.TAG_NAME, "a")
                    if "Show More" in firsta.text:
                        firsta.click()
                        all_models = all_models[1:len(all_models) - 1]
                    maketext = td.find_element(By.TAG_NAME,'b').text
                    maketext = maketext.split('\n')[0]
                    for model in all_models:
                        modeltext = model.text
                        listtoappend = [typetext,maketext,modeltext]
                        print(listtoappend)
                        finallist.append(listtoappend)
                except:
                    break
        return finallist

    def get_data_bicycle(self,finallist=[],url='https://www.fitnessrepairparts.com/type/View/5/Stationary-Bicycles'):
        self.get(url)
        typetext = "Stationary Bicycle"
        time.sleep(2)
        main_content = self.find_element(By.TAG_NAME,"tbody")
        table_rows = main_content.find_elements(By.TAG_NAME, "tr")
        for row in table_rows:
            all_table_data = row.find_elements(By.TAG_NAME, "td")
            for td in all_table_data:
                try :
                    firsta = td.find_element(By.TAG_NAME,"a")
                    all_models = td.find_elements(By.TAG_NAME, "a")
                    if "Show More" in firsta.text:
                        firsta.click()
                        all_models = all_models[1:len(all_models) - 1]
                    maketext = td.find_element(By.TAG_NAME,'b').text
                    maketext = maketext.split('\n')[0]
                    for model in all_models:
                        modeltext = model.text
                        listtoappend = [typetext,maketext,modeltext]
                        print(listtoappend)
                        finallist.append(listtoappend)
                except:
                    break
        return finallist

    def get_data_rower(self,finallist=[],url='https://www.fitnessrepairparts.com/type/View/7/Rowers'):
        self.get(url)
        typetext = "Rower"
        time.sleep(2)
        main_content = self.find_element(By.TAG_NAME,"tbody")
        table_rows = main_content.find_elements(By.TAG_NAME, "tr")
        for row in table_rows:
            all_table_data = row.find_elements(By.TAG_NAME, "td")
            for td in all_table_data:
                try :
                    firsta = td.find_element(By.TAG_NAME,"a")
                    all_models = td.find_elements(By.TAG_NAME, "a")
                    if "Show More" in firsta.text:
                        firsta.click()
                        all_models = all_models[1:len(all_models) - 1]
                    maketext = td.find_element(By.TAG_NAME,'b').text
                    maketext = maketext.split('\n')[0]
                    for model in all_models:
                        modeltext = model.text
                        listtoappend = [typetext,maketext,modeltext]
                        print(listtoappend)
                        finallist.append(listtoappend)
                except:
                    break
        return finallist

with Equipement() as bot:
    config = ConfigParser()
    config.read("config.ini")

    workbook_name = 'Final_results_of_scraping.xlsx'
    wb = Workbook()
    page = wb.active

    bot.land_first_page()
    findpartsbytype = bot.find_element(By.XPATH,"//*[contains(text(), 'Find Parts By Type')]")
    findpartsbytype.click()
    list_of_elements = bot.find_element(By.CSS_SELECTOR,"ul[style='display: block;']")

    treadmil_url = list_of_elements.find_element(By.XPATH,"//*[contains(text(), 'Treadmill Parts')]").get_attribute('href')
    eliptical_url = list_of_elements.find_element(By.XPATH,"//*[contains(text(), 'Elliptical Parts')]").get_attribute('href')
    bycicle_url = list_of_elements.find_element(By.XPATH,"//*[contains(text(), 'Stationary Bicycle Parts')]").get_attribute('href')
    rower_url = list_of_elements.find_element(By.XPATH,"//*[contains(text(), 'Rower Parts')]").get_attribute('href')

    print(treadmil_url,eliptical_url,bycicle_url,rower_url)

    result = bot.get_data_treadmill(url=treadmil_url)
    result = bot.get_data_elipticals(url=eliptical_url,finallist=result)
    result = bot.get_data_bicycle(url=bycicle_url,finallist=result)
    result = bot.get_data_rower(url=rower_url,finallist=result)

    headers = ['Type','Make','Model']
    page.append(headers)
    for row in result:
        page.append(row)
    wb.save(filename='Final_results_of_scraping.xlsx')

