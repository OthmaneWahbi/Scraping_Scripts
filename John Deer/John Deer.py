from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import undetected_chromedriver as uc
import time
from openpyxl import Workbook


if __name__ == '__main__':
    uc.TARGET_VERSION = 104
    driver = uc.Chrome()
    driver.implicitly_wait(8)
    driver.get('https://www.deere.com/en/parts-and-service/manuals-and-training/quick-reference-guides/')


    types = []
    types_links = []

    industries = driver.find_elements(By.CSS_SELECTOR,"div[data-group='qrg industries']")
    industries += driver.find_elements(By.CSS_SELECTOR,"div[data-group='qrg industries 2']")
    dont_scrape = ['Agriculture', 'Crop Care', 'OEM Engines']
    for industry in industries:
        if industry.find_element(By.TAG_NAME,"h3").text in dont_scrape:
            pass
        else:
            types.append(industry.find_element(By.TAG_NAME, "h3").text)
            types_links.append(industry.find_element(By.TAG_NAME, "a").get_attribute('href'))
    final_list=[]
    for type_text, type_link in zip(types,types_links):
        driver.get(type_link)
        time.sleep(2.5)

        sections = driver.find_elements(By.CSS_SELECTOR,"div[class='faq-section']")
        for section in sections:
            if section.find_element(By.TAG_NAME,'h3').text.lower() == 'search by model number':
                continue
            series = section.find_elements(By.CSS_SELECTOR,"div[class='faq-question']")
            for serie in series:
                serie_title = serie.find_element(By.TAG_NAME,'a')
                serie_title.send_keys(Keys.ENTER)
                time.sleep(1.5)
                serie_text = serie_title.text
                print(serie_text)
                models = serie.find_elements(By.TAG_NAME,'p')
                for model in models:
                    try:
                        model_text = model.find_element(By.TAG_NAME,'a').text
                    except:
                        continue
                    list_to_append = [type_text,'John Deere',serie_text,model_text]
                    print(list_to_append)
                    final_list.append(list_to_append)

    workbook_name = 'Final_results_of_scraping.xlsx'
    wb = Workbook()
    page = wb.active
    headers = ['Type' ,'Make' ,'Series' ,'Model' ]
    page.append(headers)
    for row in final_list:
        page.append(row)
    wb.save(filename='Final_results_of_scraping.xlsx')














