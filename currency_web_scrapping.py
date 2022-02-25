from locale import currency
from bs4 import BeautifulSoup 
from urllib.request import urlopen 

def get_currency(url):
    page_url = url

    uClient = urlopen(page_url)
    page_soup = BeautifulSoup(uClient.read(), "html.parser")
    uClient.close()

    currency_charts = page_soup.findAll("div", {"class": "cc-result"})
    currency = currency_charts[0].text
    if "AUD/HKD" in url:
        currency = currency.replace("1 AUD = ", "")
        currency = currency.replace(" HKD", "")
    elif "HKD/AUD" in url:
        currency = currency.replace("1 HKD = ", "")
        currency = currency.replace(" AUD", "")
    elif "USD/AUD" in url:
        currency = currency.replace("1 USD = ", "")
        currency = currency.replace(" AUD", "")
    currency = float(currency)
    
    return currency


