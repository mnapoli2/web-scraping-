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
    url = f'https://finance.yahoo.com/quote/{ticker_symbol}/profile'
    r = requests.get(url, headers=headers)
    soup = BeautifulSoup(r.text, 'html.parser')

    stock = {
        'stock_name': soup.find('div', {'class': 'D(ib) Mt(-5px) Maw(38%)--tab768 Maw(38%) Mend(10px) Ov(h) smartphone_Maw(85%) smartphone_Mend(0px)'}).find_all('div')[0].text.strip(),
        'Address': soup.find('div', class_='Mb(25px)').find('p', class_='D(ib) W(47.727%) Pend(40px)').get_text(separator='\n'),
        'Key Executive 1': soup.find ('table', {'class': 'W(100%)'}).find_all('td')[0].text.strip(),
        'Key Executive 2': soup.find('table', {'class': 'W(100%)'}).find_all('td')[5].text.strip(),
        'Key Executive 3': soup.find('table', {'class': 'W(100%)'}).find_all('td')[10].text.strip(),
        'Key Executive 4': soup.find('table', {'class': 'W(100%)'}).find_all('td')[15].text.strip(),
        'Key Executive 5': soup.find('table', {'class': 'W(100%)'}).find_all('td')[20].text.strip(),
        'Key Executive 6': soup.find('table', {'class': 'W(100%)'}).find_all('td')[25].text.strip(),
        'Key Executive 7': soup.find('table', {'class': 'W(100%)'}).find_all('td')[30].text.strip(),
        'Key Executive 8': soup.find('table', {'class': 'W(100%)'}).find_all('td')[35].text.strip(),
        'Key Executive 9': soup.find('table', {'class': 'W(100%)'}).find_all('td')[40].text.strip(),
        'Key Executive 10': soup.find('table', {'class': 'W(100%)'}).find_all('td')[45].text.strip(),
        'description': soup.find('p', {'class': 'Mt(15px) Lh(1.6)' }).get_text(separator='\n')
    }
    return stock

if len(sys.argv) < 2:
    print("Usage: python Homework.py <ticker_symbol1> <ticker_symbol2> ...")
    sys.exit(1)

ticker_symbols = sys.argv[1:]

stockdata = [get_data(symbol) for symbol in ticker_symbols]

with open('Napoli_stock_profile_data.json', 'w', encoding='utf-8') as f:
    json.dump(stockdata, f)

CSV_FILE_PATH = 'Napoli_stock_profile_data.csv'
with open(CSV_FILE_PATH, 'w', newline='', encoding='utf-8') as csvfile:
    fieldnames = stockdata[0].keys()
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(stockdata)

EXCEL_FILE_PATH = 'Napoli_stock_profile_data.xlsx'
df = pd.DataFrame(stockdata)
df.to_excel(EXCEL_FILE_PATH, index=False)

print('Done!')
