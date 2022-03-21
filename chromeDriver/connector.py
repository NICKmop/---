from selenium import webdriver
from bs4 import BeautifulSoup

def driver1(driver):
    driver.get('http://10.12.1.27/#/dashboard/0oKI');
    html = driver.page_source;
    return html;
def driver2(driver):
    driver.get('http://10.12.1.27/#/dashboard/woRs');
    html = driver.page_source;
    return html;
    
def connect():
    driver = webdriver.Chrome('chromeDriver\chromedriver.exe');
    driver.implicitly_wait(15);
    driver.get('http://10.12.1.27/pa/login');
    #계정정보 GET
    driver.find_element_by_name('username').send_keys('tfo-admin');
    driver.find_element_by_name('password').send_keys('123456');
    driver.find_element_by_css_selector('.md-raised.md-primary.md-button.md-default-theme').click();

    #pangData 주소
    # driver.get('http://10.12.1.27/#/dashboard/0oKI');
    # html = driver.page_source;
    
    # return 값
    soup = BeautifulSoup(driver1(driver), 'html.parser');
    soup2 = BeautifulSoup(driver2(driver), 'html.parser');
    return soup, soup2;