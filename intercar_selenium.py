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
for page_nr in range(91):
    driver.get(f"https://ro.e-cat.intercars.eu/ro/Oferta-completa/Motor-acesorii/Blocul-motorului/Volant/c/tecdoc-6200000-6211000-6211060?q=%3Adefault%3Amarket%3APKW%3AbranchAvailability%3AALL&page={page_nr}&sort=default")
    sleep(7)
    y = 300
    for timer in range(0, 8):
        driver.execute_script("window.scrollTo(0, " + str(y) + ")")
        y += 300
        sleep(1)
    page = driver.find_element(By.XPATH, "//div[@class='listing js-changeview-listwrapper is-inited']").get_attribute("innerHTML")
    html = open(f"C:\\Users\\HP\\Desktop\\ALLPARTS\\VOLANT\\intercar\\Pagina {page_nr}.html", "w", encoding="utf-8")
    html.write(page)
    html.close()



