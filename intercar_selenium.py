import os.path
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
sleep(10)
for page_nr in range(37,41):
    driver.get(f"https://ro.e-cat.intercars.eu/ro/Oferta-completa/Sistem-de-evacuare-admisie/Sistem-de-evacuare-a-gazelor/Convertor-catlitic/c/tecdoc-7200000-6010045-7213000?q=%3Adefault%3AbranchAvailability%3AALL%3AproductBrandCode%3Aicgoods_2382%3AproductBrandCode%3Aicgoods_1451%3AproductBrandCode%3Aicgoods_1477%3AproductBrandCode%3Aicgoods_1565&page={page_nr}&sort=default")
    sleep(7)
    y = 300
    for timer in range(0, 8):
        driver.execute_script("window.scrollTo(0, " + str(y) + ")")
        y += 300
        sleep(1)
    page = driver.find_element(By.XPATH, "//div[@class='listing js-changeview-listwrapper is-inited']").get_attribute("innerHTML")
    html = open(f"C:\\Users\\HP\\Desktop\\ALLPARTS\\CATALIZATOARE\\intercar_cod_made_link\\Pagina {page_nr}.html", "w", encoding="utf-8")
    html.write(page)
    html.close()



