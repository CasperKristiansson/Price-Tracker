from urllib.request import Request, urlopen
from bs4 import BeautifulSoup
import re
import pandas

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
    url = Request(
        url,
        headers={'User-Agent': 'Mozilla/5.0'}
        )
    webpage = urlopen(url).read()

    soup = BeautifulSoup(webpage, 'html.parser')
    price = str(soup.find('div', {"class": "price-value _large _center"}))
    price = re.sub("[^0-9]", "", price)
    price = 'null'                                                          #TODO Webscraping Not Working

    return price

def amazon_de_price(url):
    print("")
def amazon_sv_price(url):
    print("")

def get_data():
    product_url = []
    old_price = []
    i = 0

    excel_data_df = pandas.read_excel('Data.xlsx', sheet_name='Test')
    while i < len(excel_data_df.columns):
        column = excel_data_df.columns[i]
        data = excel_data_df[column].tolist()

        product_url.append(data[0])
        old_price.append(data[-1])
        i += 1
    
    return old_price, product_url, excel_data_df.columns.ravel()

def get_price(urls, companies):
    product_price = []
    for url in urls:
        if companies[urls.index(url)] == 'elgiganten':
            product_price.append(elgiganten_price(url))
        elif companies[urls.index(url)] == 'webbhallen':
            product_price.append(webhallen_price(url))
        elif companies[urls.index(url)] == 'netonnet':
            product_price.append(netonnet_price(url))
        elif companies[urls.index(url)] == 'amazon_de':
            product_price.append(amazon_de_price(url))
        elif companies[urls.index(url)] == 'amazon_sv':
            product_price.append(amazon_sv_price(url))
        elif companies[urls.index(url)] == '0':
            product_price.append('Error with formating in Excel (no header): {}'.format(url))
        print(product_price)
            
    return product_price

def main():
    OLD_PRICE, PRODUCT_URL, COMPANIES = get_data()
    PRODUCT_PRICE = get_price(PRODUCT_URL, COMPANIES)
    print(PRODUCT_PRICE)

if __name__ == '__main__':
    main()
