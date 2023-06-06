# coding: utf-8
from __future__ import unicode_literals

# coding: utf-8
from __future__ import unicode_literals
import uno
import requests
import asyncio
import aiohttp
import time


def _getCusips(sheet):
    cusip_list = []
    row = 6
    cell = sheet.getCellByPosition(hex(9), hex(row))
    while cell.getString() != '':
        cell_cusip = cell.getString()
        cusip_list.append(cell_cusip)
        row += 1     
        cell = sheet.getCellByPosition(hex(9), hex(row))
    return cusip_list

def _getDates(sheet):
    date_list = []
    row = 6
    cell = sheet.getCellByPosition(hex(10), hex(row))
    while cell.getString() != '':
        cell_date = cell.getString()
        date_list.append(cell_date)
        row += 1     
        cell = sheet.getCellByPosition(hex(10), hex(row))
    return date_list

"""

async def _getAPICall(cusip_date, session):
    url = f"https://www.treasurydirect.gov/TA_WS/securities/{cusip_date[0]}/{cusip_date[1]}?format=json"
    response = await session.request(method="GET", url=url)
    return response
    
async def getQuote(sheet, cusip_date, row, session):
    response = await _getAPICall(cusip_date, session)
    data = response.json()
    quote = (await data)['pricePer100']
    cell_quote = sheet.getCellByPosition(hex(12), hex(row + 5))
    cell_quote.setString(quote)
    
async def getTreasuryQuotes():
    oDoc = XSCRIPTCONTEXT.getDocument()
    sheet = oDoc.Sheets[0]
    cusips = _getCusips(sheet)
    dates = _getDates(sheet)
    cusips_dates = list(zip(cusips, dates))
    async with aiohttp.ClientSession() as session:
        await asyncio.gather(*[getQuote(sheet, cusip_date, cusips_dates.index(cusip_date), session) for cusip_date in cusips_dates])
    # await asyncio.gather(returnSymbols(sheet, symbols[0], symbols.index(symbol)), returnSymbols(sheet, symbol, symbols.index(symbol)), returnSymbols(sheet, symbol, symbols.index(symbol)

# async def getStockQuotes():
#     oDoc = XSCRIPTCONTEXT.getDocument()
#     sheet = oDoc.Sheets[0]
#     symbols = _getStockTickerSymbols(sheet)
#     await asyncio.gather(*[getQuote(sheet, symbol, symbols.index(symbol)) for symbol in symbols])

def runGetTreasuryQuotes(*args):
    asyncio.run(getTreasuryQuotes())

"""

def _getAPICall(cusip_date):
    url = f"https://www.treasurydirect.gov/TA_WS/securities/{cusip_date[0]}/{cusip_date[1]}?format=json"
    response = requests.request("GET", url)
    return response

def getQuote(sheet, cusip_date, row):
    try:
        response = _getAPICall(cusip_date)
        data = response.json()
        quote = data['pricePer100']
    except:
        quote = ''
    cell_quote = sheet.getCellByPosition(hex(14), hex(row + 6))
    cell_term = sheet.getCellByPosition(hex(11), hex(row + 6))
    cell_due = sheet.getCellByPosition(hex(12), hex(row + 6))
    cell_due.setString(data['maturityDate'][:-9])
    cell_term.setString(data['securityTerm'])
    cell_quote.setValue(quote)

def getTreasuryQuotes(*args):
    oDoc = XSCRIPTCONTEXT.getDocument()
    sheet = oDoc.Sheets[0]
    cusips = _getCusips(sheet)
    dates = _getDates(sheet)
    cusips_dates = list(zip(cusips, dates))
    for item in cusips_dates:
        getQuote(sheet, item, cusips_dates.index(item))

# Only the specified function will show in the Tools > Macro > Organize Macro dialog:
g_exportedScripts = (getTreasuryQuotes,)


