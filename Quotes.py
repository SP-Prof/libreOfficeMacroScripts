# coding: utf-8
from __future__ import unicode_literals
import requests
import uno
import asyncio
import aiohttp
import time
from com.sun.star.script.provider import XScriptContext

"""
async def count():
    print("One")
    await asyncio.sleep(1)
    print("Two")

async def asyncTest():
    
    await asyncio.gather(count(), count(), count())

def runAsyncTest():
    oDoc = XSCRIPTCONTEXT.getDocument()
    sheet = oDoc.Sheets[0]
    cell = sheet.getCellByPosition("0x03", hex(7))
    s = time.perf_counter()
    asyncio.run(asyncTest())
    elapsed = time.perf_counter() - s
    cell.setValue(elapsed)



def _getStockTickerSymbol(sheet, row):
    cell = sheet.getCellByPosition("0x03", hex(current_row))
    return cell.getString()
    
   

async def getQuote(*args):
    oDoc = XSCRIPTCONTEXT.getDocument()
    sheet = oDoc.Sheets[0]
    row = 5
    cell = sheet.getCellByPosition("0x03", hex(row))
    while cell.getString() != '':
        symbol = cell.getString()
        data = await _getAPICall(symbol)
        quote = data['regularMarketPrice']
        cell_quote = sheet.getCellByPosition("0x05", hex(row))
        cell_quote.setString(quote)
        row += 1     
        cell = sheet.getCellByPosition("0x03", hex(row))
"""
def _getStockTickerSymbols(sheet):
    symbols = ''
    row = 6
    cell = sheet.getCellByPosition("0x02", hex(row))
    while cell.getString() != '':
        cell_symbol = cell.getString()
        symbols = symbols + cell_symbol + ','
        row += 1     
        cell = sheet.getCellByPosition("0x02", hex(row))
    return symbols 

def _getAPICall(symbols):
    url = f"https://yahoo-finance15.p.rapidapi.com/api/yahoo/qu/quote/{symbols}GC=F,SI=F,BTC-USD,XMR-USD,BRL=X"
    headers = {
        "X-RapidAPI-Key": "df7437dd70msh03082390113f1dap141689jsn62c479bba512",
        "X-RapidAPI-Host": "yahoo-finance15.p.rapidapi.com"
    }
    response = requests.request("GET", url, headers=headers)
    return response

def writeCell(sheet, column, row, quote):
    cell = sheet.getCellByPosition(hex(column), hex(row))
    cell.setValue(quote)
"""
async def getQuote(sheet, symbols, row, session):
    results =  _getAPICall(symbols)
    data = results.json()[0]
    quote = data['regularMarketPrice']
    cell_quote = sheet.getCellByPosition("0x05", hex(row + 5))
    cell_quote.setString(quote)
"""
def saveQuotes(sheet, symbols):
    response = _getAPICall(symbols).json()
    results = response.json()
    results_other = results[-5:]
    stock_results = len(results) - 5
    for n in range(0, stock_results):
        quote = results[n]['regularMarketPrice']
        cell_quote = sheet.getCellByPosition("0x06", hex(n + 6))
        cell_quote.setString(quote)    
    for n in range(0, len(results_other) - 1):
        quote = results_other[n]['regularMarketPrice']
        cell_quote = sheet.getCellByPosition(hex(18), hex(n + 6))
        cell_quote.setString(quote) 
    cells_brl = sheet.getCellByPosition(hex(18), hex(13))
    cells_brl.setValue(results_other[-1]['regularMarketPrice'])
    cells_brl = sheet.getCellByPosition(hex(18), hex(14))
    cells_brl.setValue(results_other[-1]['regularMarketPrice'])


def getStockQuotes(*args):
    oDoc = XSCRIPTCONTEXT.getDocument()
    sheet = oDoc.Sheets[0]
    symbols = _getStockTickerSymbols(sheet)
    print('yes')
    saveQuotes(sheet, symbols)

# Only the specified function will show in the Tools > Macro > Organize Macro dialog:
g_exportedScripts = (getStockQuotes,)


