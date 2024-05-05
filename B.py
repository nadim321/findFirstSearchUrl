import xlrd
import xlwt
from xlwt import Workbook
from bs4 import BeautifulSoup
import requests

excel_workbook = xlrd.open_workbook('Company_info_1946.xls')
# excel_worksheet_2020 = excel_workbook.sheet_by_index(0)
excel_worksheet_2021 = excel_workbook.sheet_by_name('Company_info_1946')
url = 'https://www.google.com/search'

headers = {
	'Accept' : '*/*',
	'Accept-Language': 'en-US,en;q=0.5',
	'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.82',
}
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')
for x in range(0, 1846, 1):
    companyName = excel_worksheet_2021.cell_value(x, 0);
    companyCountry = excel_worksheet_2021.cell_value(x, 1);
    print(x)

    sheet1.write(x, 0, companyName)
    sheet1.write(x, 1, companyCountry)

    parameters = {'q': companyName+' '+companyCountry}

    content = requests.get(url, headers=headers, params=parameters).text
    soup = BeautifulSoup(content, 'html.parser')

    search = soup.find(id='search')
    first_link = search.find('a')
    # print( companyName + ' ' + companyCountry + ' ' + first_link['href'])
    sheet1.write(x, 2, first_link['href'])

wb.save('Company_info_1946_update.xls')