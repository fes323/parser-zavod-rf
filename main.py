import re
import requests
from bs4 import BeautifulSoup
import openpyxl
import io
from lxml import etree

# Функция для получения всех ссылок на компании с одной страницы
def get_company_links(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    company_links = [a['href'] for a in soup.select('.factory__title')]
    return company_links

# Функция для получения информации о компании
def get_company_info(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    xpath_soup = BeautifulSoup(response.content, 'html.parser')
    dom = etree.HTML(str(xpath_soup))

    company_name = dom.xpath('/html/body/div[1]/div/div[1]/div[2]/div/div[2]/div/h1')[0].text

    address = dom.xpath('//*[@id="contact-company"]/ul/li[1]/span[2]')[0].text

    phone_elements = soup.select('.content-list__descr a[href^="tel:"]')
    phones = [phone['href'].replace('tel:', '') for phone in phone_elements] if phone_elements else []

    email_elements = soup.select('.content-list__descr a[href^="mailto:"]')
    emails = [email['href'].replace('mailto:', '') for email in email_elements] if email_elements else []

    website_element = soup.select_one('.content-list__descr a[target="_blank"]')
    website = website_element['href'] if website_element else ''

    return company_name, address, phones, emails, website

# Функция для сохранения данных в .xlsx файл
def save_to_excel(data, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = ['Ссылка', 'Наименование компании', 'Адрес', 'Телефоны', 'Почты', 'Сайт']
    ws.append(headers)
    for row in data:
        # Преобразование списков в строки, разделенные запятыми
        row[3] = ', '.join(row[3])  # Телефоны
        row[4] = ', '.join(row[4])  # Почты
        ws.append(row)

    # Сохраняем данные в байтовом формате с помощью io.BytesIO
    with io.BytesIO() as f:
        wb.save(f)
        f.seek(0)
        with open(filename, 'wb') as excel_file:
            excel_file.write(f.read())

if __name__ == "__main__":
    base_url = "https://xn--80aegj1b5e.xn--p1ai/factories/kotelnye-zavody"
    all_company_links = []
    page = 0

    # Собираем все ссылки на компании со всех страниц
    while True:
        current_page_url = f"{base_url}?page={page}"
        company_links_on_page = get_company_links(current_page_url)
        if not company_links_on_page:
            break
        all_company_links.extend(company_links_on_page)
        page += 1

    # Собираем информацию о каждой компании
    all_company_info = []
    for link in all_company_links:
        company_url = f"https://xn--80aegj1b5e.xn--p1ai{link}"
        company_info = get_company_info(company_url)
        all_company_info.append([company_url] + list(company_info))

    # Сохраняем данные в .xlsx файл
    save_to_excel(all_company_info, "kotelnye-zavody.xlsx")

    print("Данные успешно записаны в файл companies_info.xlsx")
