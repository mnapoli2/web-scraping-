import json
import csv
import sys
from typing import Any, Dict
import requests
from bs4 import BeautifulSoup
import pandas as pd


def get_data(ticker_symbol):
    print('Getting stock data of ', ticker_symbol)
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'}
    url = f'https://finance.yahoo.com/quote/{ticker_symbol}/holders'
    r = requests.get(url, headers=headers)
    soup = BeautifulSoup(r.text, 'html.parser')

    stock = {
        'stock_name': soup.find('div', { 'class': 'D(ib) Mt(-5px) Maw(38%)--tab768 Maw(38%) Mend(10px) Ov(h) smartphone_Maw(85%) smartphone_Mend(0px)'}).find_all('div')[0].text.strip(),
        'Major holder 1' : soup.find('div', { 'class': 'W(100%) Mb(20px)'}).find_all('tr')[0].text.strip(),
        'Major holder 2': soup.find('div', {'class': 'W(100%) Mb(20px)'}).find_all('tr')[1].text.strip(),
        'Major holder 3': soup.find('div', {'class': 'W(100%) Mb(20px)'}).find_all('tr')[2].text.strip(),
        'Major holder 4': soup.find('div', {'class': 'W(100%) Mb(20px)'}).find_all('tr')[3].text.strip(),
        'Top Institutional Holder 1' : soup.find('div', {'class' : 'Mt(25px) Ovx(a) W(100%)'}).find_all('td')[0].text.strip(),
        'Top Institutional Holder 2': soup.find('div', {'class': 'Mt(25px) Ovx(a) W(100%)'}).find_all('td')[5].text.strip(),
        'Top Institutional Holder 3': soup.find('div', {'class': 'Mt(25px) Ovx(a) W(100%)'}).find_all('td')[10].text.strip(),
        'Top Institutional Holder 4': soup.find('div', {'class': 'Mt(25px) Ovx(a) W(100%)'}).find_all('td')[15].text.strip(),
        'Top Institutional Holder 5': soup.find('div', {'class': 'Mt(25px) Ovx(a) W(100%)'}).find_all('td')[20].text.strip(),
        'Top Institutional Holder 6': soup.find('div', {'class': 'Mt(25px) Ovx(a) W(100%)'}).find_all('td')[25].text.strip(),
        'Top Institutional Holder 7': soup.find('div', {'class': 'Mt(25px) Ovx(a) W(100%)'}).find_all('td')[30].text.strip(),
        'Top Institutional Holder 8': soup.find('div', {'class': 'Mt(25px) Ovx(a) W(100%)'}).find_all('td')[35].text.strip(),
        'Top Institutional Holder 9': soup.find('div', {'class': 'Mt(25px) Ovx(a) W(100%)'}).find_all('td')[40].text.strip(),
        'Top Institutional Holder 10': soup.find('div', {'class': 'Mt(25px) Ovx(a) W(100%)'}).find_all('td')[45].text.strip()

    }

    return stock

if len(sys.argv) < 2:
    print("Usage: python napoli_scraper_holder.py <ticker_symbol1> <ticker_symbol2> ...")
    sys.exit(1)

ticker_symbols = sys.argv[1:]

stockdata = [get_data(symbol) for symbol in ticker_symbols]

with open('napoli_stock_holder_data.json', 'w', encoding='utf-8') as f:
    json.dump(stockdata, f)

CSV_FILE_PATH = 'napoli_stock_holder_data.csv'
with open(CSV_FILE_PATH, 'w', newline='', encoding='utf-8') as csvfile:
    fieldnames = stockdata[0].keys()
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(stockdata)

EXCEL_FILE_PATH = 'napoli_stock_holder_data.xlsx'
df = pd.DataFrame(stockdata)
df.to_excel(EXCEL_FILE_PATH, index=False)

print('Done!')
