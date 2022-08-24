import os.path

import selenium
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from openpyxl import load_workbook

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://ro.e-cat.intercars.eu")
sleep(5)
driver.find_element(By.XPATH, "//input[@class='form-control form-control  bf-required']").send_keys("emanuel_b1998@yahoo.com")
driver.find_element(By.XPATH, "//input[@class='form-control form-control']").send_keys("Hala@Bilca2021")
sleep(1)
driver.find_element(By.XPATH, "//button[@class='btn btn-default btn col-sm-12']").click()
sleep(5)
driver.get("https://ro.e-cat.intercars.eu/ro/")
sleep(5)
workbook = load_workbook("C:\\Users\\Gh0sT\\Desktop\\ALLPARTS\\VOLANT\\cod_producator_link.xlsx")
worksheet = workbook["Sheet1"]
column_link = worksheet["C"]
intercar_link = [column_link[x].value for x in range(len(column_link))]
column_code = worksheet["A"]
code_list = [column_code[x].value for x in range(len(column_code))]
for link in intercar_link[1:]:
    code = code_list[intercar_link.index(link)]
    if os.path.isfile(f"C:\\Users\\Gh0sT\\Desktop\\ALLPARTS\\VOLANT\\intercar_pret_livrare_descr\\{code}.html"):
        print(f"Produs extractat: {code}")
    else:
        print(f"Produs in proces: {code}")
        driver.get(link)
        sleep(5)
        html_parameters = None
        html_other_numbers = None
        html_applications = None
        for tab in driver.find_elements(By.XPATH, "//div[@class='tabs__item']"):
            if tab.text == "ECHIVALENTE":
                tab.click()
                sleep(2)
            if "ALTE" in tab.text:
                tab.click()
                sleep(2)
        for tab in driver.find_elements(By.XPATH, "//div[@class='tabs__item']"):
            if tab.text == "APLICATII":
                tab.click()
                sleep(2)
                for branch in driver.find_elements(By.XPATH, "//div[@class='tree__branch']"):
                    branch.click()
                sleep(2)
                for leaf in driver.find_elements(By.XPATH, "//div[@class='tree__leaf js-tree-trigger']"):
                    leaf.click()
                sleep(2)
                html_applications = driver.find_element(By.XPATH, "//div[@class='layoutproductdetails__tabs layoutproductdetails__tabs--doublerow productprice--productdetails productretailprice--productdetails']").get_attribute("innerHTML")
                price = driver.find_element(By.XPATH, "//div[@class='buybox js-onboarding-productdetails-buybox buybox--']").get_attribute("innerHTML")
                try:
                    image = driver.find_element(By.XPATH, "//div[@class='productcarousel__mainitem slick-slide slick-current slick-active']").get_attribute("innerHTML")
                except:
                    image = "Fara imagine"
                descriere_tehnica = driver.find_element(By.XPATH, "//div[@class='productinfo__technicaldesc']").get_attribute("innerHTML")
                tip_produs = driver.find_element(By.XPATH, "//div[@class=' productname productname--productinfo']").get_attribute("innerHTML")
        html = open(f"C:\\Users\\Gh0sT\\Desktop\\ALLPARTS\\VOLANT\\intercar_pret_livrare_descr\\{code}.html", "w", encoding="utf-8")
        html.write(tip_produs)
        html.write(descriere_tehnica)
        html.write(image)
        html.write(html_applications)
        html.write(price)
        html.close()
        sleep(2)