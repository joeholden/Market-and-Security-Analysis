import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import pyexcel as p
import re


HEADERS = {
    'a': 'date',
    'b': 'NYSE Advancing Issues',
    'c': 'NYSE Declining issues',
    'd': 'NYSE up volume',
    'e': 'NYSE down volume',
    'f': 'DJIA Close',
    'g': 'NYSE A-D',
    'h': '10% Trend A-D',
    'i': '5% Trend A-D',
    'j': 'McC A-D Osc',
    'k': 'McC Summation Index',
    'l': 'Osc unchanged tomorrow',
    'm': 'Osc to zero tomorrow',
    'n': 'NYSE UV-DV',
    'o': '10% Trend UV-DV',
    'p': '5% Trend UV-DV',
    'q': 'McC UV-DV Osc',
    'r': 'McC Vol Summation Index',
    's': 'UV-DV for unchanged osc tomorrow',
    't': 'UV-DV for osc to zero tomorrow',
    'u': '10% trend',
    'v': '5% trend',
    'w': 'price oscillator',
    'x': 'price for unchanged oscillator',
}


def download_mclellan_website_data():
    url = 'https://www.mcoscillator.com/market_breadth_data/'
    s = Service(ChromeDriverManager().install())
    op = webdriver.ChromeOptions()
    op.add_argument('headless')
    driver = webdriver.Chrome(service=s)

    driver.get(url)
    download_link = driver.find_element(by=By.XPATH, value='//*[@id="data_table"]/a[1]/img')
    download_link.click()

    time.sleep(1)


# alter this so you can pre-process the sheet in openpyxl to get correct format
def clean_excel_data():
    """Takes the xls file that was downloaded from download_mclellan_website_data, converts it to .xlsx
    and then uses openpyxl to process it to a cleaned dataframe.
    """
    global first_row_to_keep, first_row_to_delete
    p.save_book_as(file_name='C:/Users/joema/Downloads/OSC-DATA.xls',
                   dest_file_name='C:/Users/joema/Downloads/OSC-DATA.xlsx')
    wb = openpyxl.load_workbook('C:/Users/joema/Downloads/OSC-DATA.xlsx')
    ws = wb['A']

    # Get the correct bounds of the dataframe
    # the group(0) notation just accesses the text of the match object
    # Checks first 30 rows for a date-like object

    for cell_row in range(1, 30):
        c = ws[f'A{cell_row}']
        cell_contents = str(c.value).split(" ")[0]
        try:
            re.match(r"^([\d]{4})-([\d]+)-([\d]+)$", cell_contents).group(0)
            first_row_to_keep = cell_row
            break
        except AttributeError:
            pass
    # first row to keep: most likely this will be cell A9 unless the format changes

    # Check for first row to delete
    # if cell.value: just checks if the cell exists. In openpyxl cells are created as they are populated
    # They do not exist otherwise

    for cell_row in range(first_row_to_keep, 500):
        if not ws[f'B{cell_row}'].value:
            first_row_to_delete = cell_row
            break

    # delete rows to clean
    ws.delete_rows(1, first_row_to_keep - 1)
    ws.delete_rows(first_row_to_delete - first_row_to_keep + 1, 500)
    wb.save('cleaned.xlsx')


def get_advance_decline_data():
    """gets a plot of the advance decline line (10 day sma)"""
    data = pd.read_excel('cleaned.xlsx', engine='openpyxl', names=HEADERS.values(), )
    data['10 day sma'] = data.iloc[:, 1].rolling(window=10).mean()

    advance_decline = data['NYSE A-D']
    date = data['date']
    sma10 = data['10 day sma']

    plt.plot(date, sma10)
    plt.show()


def get_volume_data():
    """Gets a plot of the uvol - dvol on the NYSE"""
    data = pd.read_excel('cleaned.xlsx', engine='openpyxl', names=HEADERS.values())
    data['10 day sma'] = data.iloc[:, 13].rolling(window=10).mean()

    uvol_dvol = data['NYSE UV-DV']
    date = data['date']
    sma10 = data['10 day sma']

    plt.plot(date, sma10, color='purple')
    plt.show()


download_mclellan_website_data()
clean_excel_data()
get_advance_decline_data()
get_volume_data()

