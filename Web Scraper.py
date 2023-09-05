import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta

#This is the Url from google finance for Nike's Stock prices
url = 'https://www.google.com/finance/quote/NKE:NYSE?hl=en'

#Sends request to the server to retrive data
page = requests.get(url)
soup = BeautifulSoup(page.text, 'html.parser')

#specifying the specific data value that we want to scrape
price_div = soup.find('div', {'class': ['YMlKec fxKbKc']})

#skips the first character ($)
price_text = price_div.text[1:]

#Loading in our excel file
workbook = load_workbook('Project Data.xlsx')
sheet = workbook.active

#getting current day to update values monthly by cell
current_date = datetime.now().date()
updated_cell = f'A{current_date.day}'

#updates the value in the respective cell to the retrived value
sheet[updated_cell] = price_div.text

#saves the spreadsheet
workbook.save('Project Data.xlsx')