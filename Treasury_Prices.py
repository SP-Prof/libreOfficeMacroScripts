import requests
import bs4
import pandas as pd
                      
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

class PriceTable:
    def __init__(self):
        self.url ='https://www.federalinvestments.gov/GA-FI/FedInvest/todaySecurityPriceDetail'
        self.bonds_dict = {}
    def request_table(self):
        res = requests.request("GET", self.url, verify=False) 
        return res
    def get_rows(self):
        res = self.request_table()
        soup = bs4.BeautifulSoup(res.text, 'html.parser')
        table = soup.find("table", class_='data1')
        rows = table.find_all('tr')
        return rows
    def build_dict(self):
        rows = self.get_rows()
        for row in rows[1:]:
            entry = row.find_all('td')
            self.bonds_dict[entry[0].text] = entry[1:]
    """
    for row in rows[1:]:    
        columns = []
        for col in row.find_all('td'):
            columns.append(col)
        cusip = columns[0].text
        security_type = columns[1].text
        rate = columns[2].text
        maturity_date = columns[3].text
        call_date = columns[4].text
        buy = float(columns[5].text)
        sell = float(columns[6].text)
        end_of_day = columns[7].text
        df = df._append({'CUSIP': cusip,  'SECURITY TYPE': security_type, 'RATE': rate, 'MATURITY DATE': maturity_date, 
                        'CALL DATE': call_date, 'BUY': buy, 'SELL': sell, 'END OF DAY': end_of_day}, ignore_index=True)
    """


def getTreasuryQuotes(*args):
    oDoc = XSCRIPTCONTEXT.getDocument()
    sheet = oDoc.Sheets[0]
    cusips = _getCusips(sheet)
    bonds = PriceTable()
    bonds.build_dict()
    prices = []
    for item in cusips:
        low = float(bonds.bonds_dict[item][4].text)
        high = float(bonds.bonds_dict[item][5])
        average = (low + high)/2
        prices.append(average)
    row = 6
    cell = sheet.getCellByPosition(hex(14), hex(row))
    for price in prices:
        cell.setValue(price)
        cell = sheet.getCellByPosition(hex(9), hex(row))
        row += 1
    
