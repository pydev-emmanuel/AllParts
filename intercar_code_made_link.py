import os
import xlsxwriter
from openpyxl import load_workbook
from bs4 import BeautifulSoup as soup
from time import sleep


directory = "C:\\Users\\HP\\Desktop\\ALLPARTS\\FILTRE\\AER"
produs = []
producator = None
cod_piesa = None
link_piesa = None
for html_file in os.listdir(f"{directory}\\intercar_cod_made_link"):
    print(html_file)
    path = f"{directory}\\intercar_cod_made_link\\{html_file}"
    html = open(path, mode="r", encoding="utf-8")
    content = html.read()
    bs_contents = soup(content, "lxml")
    for product in bs_contents.find_all("tr", class_="listingcollapsed__content js-quantity-wrapper"):
        try:
            cod_piesa = product.find("div", class_="listingcollapsed__activenumbercontainer").a.text
            cod_piesa_list = list(cod_piesa)
            for i in cod_piesa_list:
                if i == "\n":
                    cod_piesa_list.remove(i)
            cod_piesa = "".join(cod_piesa_list).replace(" ", "")
        except TypeError:
            pass
        try:
            producator = product.find("div", class_="listingcollapsed__manufacturer").findNext("img")["title"]
        except TypeError:
            pass
        try:
            link_piesa = f'https://ro.e-cat.intercars.eu/{product.find("div", class_="listingcollapsed__activenumbercontainer").a["href"]}'
        except TypeError:
            pass
        print(cod_piesa, producator, link_piesa)
        produs.append([cod_piesa, producator, link_piesa])

workbook = xlsxwriter.Workbook(f"{directory}\\cod_producator_link.xlsx")
worksheet = workbook.add_worksheet("Sheet1")
column_cod_piesa = 0
column_producator = 1
column_link_piesa = 2
row = 0
for item in produs:
        worksheet.write(row, column_cod_piesa, item[0])
        worksheet.write(row, column_producator, item[1])
        worksheet.write(row, column_link_piesa, item[2])
        row += 1
workbook.close()
