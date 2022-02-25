from asyncio.windows_events import NULL
from email import header
from lib2to3.pgen2.pgen import ParserGenerator
import requests
from bs4 import BeautifulSoup

def create_url(symbol):
    url = f'https://finance.yahoo.com/quote/{symbol}'
    return url

def get_html(url):
    header = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.190 Safari/537.36'}
    response = requests.get(url, headers = header)

    if response.status_code == 200:
        return response.text
    else:
        return None

def parse_data(html, symbol):
    page_soup = BeautifulSoup(html,'html.parser')

    current_price = float(page_soup.find('fin-streamer', {'class': 'Fw(b) Fz(36px) Mb(-4px) D(ib)'}).text)
    return current_price

def stock(code):
    url = create_url(code)
    # get html
    html = get_html(url)
    # parse data
    price = parse_data(html, stock)

    return price