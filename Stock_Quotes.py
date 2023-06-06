# coding: utf-8
from __future__ import unicode_literals
import requests

def _getStockTickerSymbols(sheet):
    symbols = ''
    row = 5
    cell = sheet.getCellByPosition("0x03", hex(row))
    while cell.getString() != '':
        cell_symbol = cell.getString()
        symbols = symbols + cell_symbol + ','
        row += 1     
        cell = sheet.getCellByPosition("0x03", hex(row))
    symbols = symbols[:-1]
    return symbols 

def _getAPICall(symbols):
    url = f"https://yahoo-finance15.p.rapidapi.com/api/yahoo/qu/quote/{symbols}"
    headers = {
        "X-RapidAPI-Key": "df7437dd70msh03082390113f1dap141689jsn62c479bba512",
        "X-RapidAPI-Host": "yahoo-finance15.p.rapidapi.com"
    }
    # response = await requests.request("GET", url, headers=headers, params=querystring)
    response = requests.request("GET", url=url, headers=headers)
    return response

def saveQuotes(sheet, symbols):
    results = _getAPICall(symbols).json()
    for n in range(0, len(results)):
        quote = results[n]['regularMarketPrice']
        cell_quote = sheet.getCellByPosition("0x05", hex(n + 5))
        cell_quote.setString(quote)    

def getStockQuotes():
    oDoc = XSCRIPTCONTEXT.getDocument()
    sheet = oDoc.Sheets[0]
    symbols = _getStockTickerSymbols(sheet)
    saveQuotes(sheet, symbols)
    

g_exportedScripts = (getStockQuotes,)