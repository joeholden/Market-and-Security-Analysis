from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import datetime

"""
This script scrapes CBOE website for total and equity put/call ratios and updates a central excel file.
This file will be used to create a put call oscillator plot
"""


def download_put_call():
    url = 'https://www.cboe.com/us/options/market_statistics/daily/'
    s = Service(ChromeDriverManager().install())
    op = webdriver.ChromeOptions()
    op.headless = True
    driver = webdriver.Chrome(service=s, options=op)

    driver.get(url)
    total_put_call = driver.find_element(by=By.XPATH,
                                         value='/html/body/main/section/div/div[2]/table/tbody/tr[1]/td[2]')
    equity_put_call = driver.find_element(by=By.XPATH,
                                          value='/html/body/main/section/div/div[2]/table/tbody/tr[4]/td[2]')
    return total_put_call.text, equity_put_call.text


def update_excel_sheet():
    headers = ['Day', 'Total Put/Call', 'Equity Put/Call']
    wb = load_workbook('C:/Users/joema/PycharmProjects/stocks/excel files/put_call.xlsx')
    ws = wb.active

    # Get First Empty Row to Append @
    for row in range(1, 10000):
        cell = ws[f'A{row}']
        if cell.value is None:
            first_empty_row = row
            break

    today = datetime.date.today()
    total_pc, equity_pc = download_put_call()

    ws[f'A{first_empty_row}'], ws[f'B{first_empty_row}'], ws[f'C{first_empty_row}'] = today, total_pc, equity_pc
    wb.save('C:/Users/joema/PycharmProjects/stocks/excel files/put_call.xlsx')


# Monday is 0, Sunday is 6
if datetime.date.today().weekday() in [0, 1, 2, 3, 4, 5, 6]:
    update_excel_sheet()
