#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math

#saving the stocks as a pandas data frame
stocks = pd.read_csv('sp_500_stocks.csv')

#importing API token
from api_token import IEX_CLOUD_API_TOKEN

#column names for Excel sheet
data_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']


symbol_list = []
min = 0
max = 100

#Creating sublists of Stocks (0-100, 101-200, 201-300, 301-400, 401-500, 501-505)
#as batch Api call can get 100 stocks in one call
for i in range(6):
    symbols = []
    for z in range(min, max):
        symbols.append(stocks['Ticker'][z])
    min = min + 100
    if (min == 500) :
      max = 505
    else:
        
      max = max + 100
    symbol_list.append(symbols)
    
symbol_strings = []    
 
for symbols in symbol_list:
        symbol_strings.append(",".join(symbols))
        
final_DataFrame = pd.DataFrame(columns = data_columns)

#fetching stocks batch by batch
for symbols in symbol_strings:
    batch_api_call = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbols}&types=quote&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call).json()
#     print(data)
    
    for sym in symbols.split(','):
        
        #adding stocks and their data to pandas data frame
        final_DataFrame = final_DataFrame.append(
        pd.Series(
        [
           sym,
           data[sym]['quote']['latestPrice'] ,
           data[sym]['quote']['marketCap'],
            'N/A'
        ],index = data_columns),
        ignore_index=True)

#Mathematical calculations to find out how many shares to buy for each stock
value = input("Please enter your portfolio size: ")

try:
    portfolio_size = float(value)
except:
    value = input("Please input a number: ")
    while isinstance(value, str):
        value = input("Please input a number: ")
    portfolio_size = float(value)

position_size = portfolio_size/len(final_DataFrame)

for index, row in final_DataFrame.iterrows():
    final_DataFrame.loc[index, 'Number of Shares to Buy'] = math.floor(position_size/row['Stock Price'])

    
#Initiating Excel writer object
writer = pd.ExcelWriter('recommended trades.xlsx', engine = 'xlsxwriter')
final_DataFrame.to_excel(writer, 'Recommended Trades', index = False)    

#Setting layout for Excel sheet
background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
{
    'font_color': font_color, 
    'bg_color': background_color,
    'border': 1
})

dollar_format = writer.book.add_format(
{
    'num_format': '$0.00',
    'font_color': font_color, 
    'bg_color': background_color,
    'border': 1
})

integer_format = writer.book.add_format(
{
    'num_format': '0',
    'font_color': font_color, 
    'bg_color': background_color,
    'border': 1
})


# writer.sheets['Recommended Trades'].set_column('A:A', 18, string_format)
# writer.sheets['Recommended Trades'].set_column('B:B', 18, dollar_format)
# writer.sheets['Recommended Trades'].set_column('C:C', 18, integer_format)
# writer.sheets['Recommended Trades'].set_column('D:D', 18, integer_format)
# writer.save()

# writer.sheets['Recommended Trades'].write('A1', 'Ticker', string_format)
# writer.sheets['Recommended Trades'].write('B1', 'Stock Price', string_format)
# writer.sheets['Recommended Trades'].write('C1', 'Market Capitalization', string_format)
# writer.sheets['Recommended Trades'].write('D1', 'Number of Shares to Buy', string_format)

column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C': ['Market Capitalization', dollar_format],
    'D': ['Number of Shares to Buy', integer_format]
}

#loop to format headers and all cells in one go
for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)

#saving Excel Sheet
writer.save()

        

