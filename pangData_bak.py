from ast import Try
from cmath import e
from warnings import catch_warnings
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import openpyxl

wb = openpyxl.Workbook();
sheet1 = wb['Sheet'];

driver = webdriver.Chrome('chromedriver.exe');
driver.implicitly_wait(15);

driver.get('http://10.12.1.27/pa/login');

driver.find_element_by_name('username').send_keys('tfo-admin');
driver.find_element_by_name('password').send_keys('123456');

driver.find_element_by_css_selector('.md-raised.md-primary.md-button.md-default-theme').click();

driver.get('http://10.12.1.27/#/dashboard/0oKI');

html = driver.page_source;
soup = BeautifulSoup(html, 'html.parser');

#텍스트 값에 변수 등록 필요
listBox = [];
listTextBox = [];
ulTotal = soup.select_one('#dashboardContentDiv > span > span > ul');
liTotal = ulTotal.select('li');

for i in liTotal:
    listBox.append(i);

for j in range(1,len(listBox)):
    liText = listBox[j].text;
    # print("Text Type : ", type(liText));
    # sheet1.cell(row=j, column=j).value = listBox;
    listTextBox.append(liText);

print(type(listTextBox));
try:
    # sheet1['A1'] = listTextBox;
    sheet1.append(listTextBox);
    wb.save('test.xlsx');
except:
    print('e');

