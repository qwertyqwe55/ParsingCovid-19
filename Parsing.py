
from selenium import webdriver
import requests
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime
import openpyxl
URL = 'https://covid19.who.int/table'

HEADERS = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.106 YaBrowser/21.6.0.620 Yowser/2.5 Safari/537.36',
    'accept': '*/*'}

FILE = 'Coronavirus.csv'

countries = ['Global', 'Russian Federation']


def get_html(url, params=None):
    r = requests.get(url, headers=HEADERS, params=params)
    return r


def get_content():
    driver = webdriver.Chrome()
    driver.get(URL)
    try:
        element = WebDriverWait(driver, 40).until(
            EC.presence_of_element_located((By.CLASS_NAME, "tbody")))
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        world = soup.find_all('div',class_ = "sc-fznYue iNYtic th")
        countries = soup.find_all('div', class_='tr depth_0')
    finally:
        driver.quit()

    list = []

    list.append({
        'country':'Global',
        'cases': world[1].get_text().replace("\\xa", " "),
        'death':world[3].get_text().replace("\\xa", " ")
    })

    for country in countries:
        if country.find('div', class_ = 'sc-AxjAm sc-fzoMdx cTipgc').get_text() == "Russian Federation":
            list.append({
                'country': 'Russian Federation',
                'cases' : country.find('div', class_ = 'column_Confirmed td').get_text(),
                'death' : country.find('div', class_ = 'column_Deaths td').get_text()
            })
    return list
def change_line(list):
    for element in list:
        s = element.get('cases').replace('\\xa',"")
        print(element.get('cases'))
        element['cases'] = s
        element['death'] = element.get('death').replace('\\xa'," ")

    print(list)

def write_in_file(list, file_name):
    wb = openpyxl.load_workbook(file_name)
    now = datetime.datetime.now()
    info = [now.strftime('%d.%m.%Y'),list[0].get('cases'), list[0].get('death'), list[1]['cases'], list[1]['death']]

    page = wb.active

    page.append(info)
    wb.save(filename=file_name)



def parse():
    list = get_content()
    file_name = input()
    write_in_file(list,file_name)


parse()
