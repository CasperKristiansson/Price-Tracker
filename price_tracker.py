from urllib.request import Request, urlopen
import re
from bs4 import BeautifulSoup
import pandas
import requests
from pyexcel_xlsx import get_data, save_data

def elgiganten_price(url):
    url = Request(
        url,
        headers={'User-Agent': 'Mozilla/5.0'}
        )
    webpage = urlopen(url).read()

    soup = BeautifulSoup(webpage, 'html.parser')
    price = str(soup.find('div', {"class": "product-price-container"}))
    price = re.sub("[^0-9]", "", price)

    return price

def netonnet_price(url):
    url = Request(
        url,
        headers={'User-Agent': 'Mozilla/5.0'}
        )
    webpage = urlopen(url).read()

    soup = BeautifulSoup(webpage, 'html.parser')
    price = str(soup.find('div', {"class": "price-big"}))
    price = re.sub("[^0-9]", "", price)

    return price

def webhallen_price(url):
    with requests.Session() as session:
        response = session.get(url).json()
        price = response["product"]["price"]["price"]

    return price

def amazon_de_price(url):
    headers = {
        "User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.90 Safari/537.36',
        "Accept-Encoding":"gzip, deflate",
        "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "DNT":"1",
        "Connection":"close",
        "Upgrade-Insecure-Requests":"1"
    }

    page = requests.get(url, headers=headers)
    soup = BeautifulSoup(page.content, 'html.parser')

    price = soup.find(id="priceblock_ourprice").get_text()
    price = price[ 0 : price.index(".")]
    price = re.sub("[^0-9]", "", price)
    price = round(int(price) * 10.15)
    return price

def amazon_sv_price(url):
    headers = {
        "User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.90 Safari/537.36',
        "Accept-Encoding":"gzip, deflate",
        "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "DNT":"1",
        "Connection":"close",
        "Upgrade-Insecure-Requests":"1"
    }

    page = requests.get(url, headers=headers)
    soup = BeautifulSoup(page.content, 'html.parser')

    price = soup.find(id="priceblock_ourprice").get_text()
    price = price[ 0 : price.index(",")]
    price = re.sub("[^0-9]", "", price)

    return price

def data():
    product_url = []
    i = 0

    excel_data_df = pandas.read_excel('Data.xlsx', sheet_name='Sheet1')
    while i < len(excel_data_df.columns):
        column = excel_data_df.columns[i]
        selected_data = excel_data_df[column].tolist()

        product_url.append(selected_data[0])

        i += 1

    return product_url, excel_data_df.columns.ravel(), excel_data_df

def get_price(urls, companies):
    product_price = []
    for url in urls:
        if companies[urls.index(url)] == 'elgiganten':
            try:
                product_price.append(elgiganten_price(url))
            except:
                product_price.append("null")
        elif companies[urls.index(url)] == 'webbhallen':
            #try:
            product_price.append(webhallen_price(url))
            #except:
                #product_price.append("null")
        elif companies[urls.index(url)] == 'netonnet':
            try:
                product_price.append(netonnet_price(url))
            except:
                product_price.append("null")
        elif companies[urls.index(url)] == 'amazon_de':
            try:
                product_price.append(amazon_de_price(url))
            except:
                product_price.append("null")
        elif companies[urls.index(url)] == 'amazon_sv':
            try:
                product_price.append(amazon_sv_price(url))
            except:
                product_price.append("null")
        elif companies[urls.index(url)] == '0':
            product_price.append('Error with formating in Excel (no header): {}'.format(url))

    return product_price

def write_data(product_price, excel_data, companies):

    rows, columns = excel_data.shape
    j = 0
    i = 0
    new_data = []

    new_data.append([])
    nested_list = new_data[0]
    for companies_name in companies:
        nested_list.append(companies_name)

    while i < rows:
        new_data.append([])
        nested_list = new_data[i+1]
        while j < columns:
            column = excel_data.columns[j]
            selected_data = excel_data[column].tolist()

            nested_list.append(selected_data[i])
            j+=1
        j=0
        i+=1

    new_data.append([])
    nested_list = new_data[-1]
    for item in product_price:
        nested_list.append(item)

    data = get_data("Data.xlsx")
    data.update({"Sheet1": new_data})
    save_data("Data.xlsx", data)

def main():
    PRODUCT_URL, COMPANIES, EXCEL_DATA = data()
    PRODUCT_PRICE = get_price(PRODUCT_URL, COMPANIES)
    write_data(PRODUCT_PRICE, EXCEL_DATA, COMPANIES)

if __name__ == '__main__':
    main()
