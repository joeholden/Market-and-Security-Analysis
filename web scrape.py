from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC


s = Service(ChromeDriverManager().install())
op = webdriver.ChromeOptions()
# op.add_argument('headless')
driver = webdriver.Chrome(service=s)


def get_last_price(ticker):
    driver.get(f"https://www.nasdaq.com/market-activity/stocks/{ticker}")
    return WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[@class='symbol-page-header__pricing-price' and text()]"))).text


print(get_last_price('MA'))

