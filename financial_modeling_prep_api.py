import requests
import os

# noinspection SpellCheckingInspection
portfolio = {
    'AAPL': 46.78,
    'ADBE': 2,
    'AMAT': 16.031,
    'AMD': 85,
    'AMZN': 1.076,
    'F': 205.694,
    'FB': 3,
    'GS': 1,
    'GOOGL': 2,
    'HON': 2.018,
    'MA': 2.003,
    'MRVL': 27.084,
    'MS': 17.145,
    'MSFT': 13.179,
    'NUE': 9.032,
    'NVDA': 15.01,
    'PYPL': 2,
    'SBUX': 5,
    'QCOM': 4
}
CASH = 2645

# portfolio = {
#     'AMD': 100,
# }


def get_closing_price_and_volume(tickers):
    data = []
    for company in tickers:
        response = requests.get(f'https://financialmodelingprep.com/api/v3/quote-short/{company}?apikey={os.environ["APIKEY"]}')
        data.append(response.json())
    return data


def get_portfolio_value(portfolio_data):
    portfolio_value = CASH
    for datapoint in portfolio_data:
        company = datapoint[0]['symbol']
        closing_price = datapoint[0]['price']
        number_of_shares = portfolio[company]
        position_value = number_of_shares * closing_price
        portfolio_value += position_value
    portfolio_value_formatted = "${:,.2f}".format(portfolio_value)
    return portfolio_value_formatted


def get_historical_closing_prices(company):
    response = requests.get(f'https://financialmodelingprep.com/api/v3/historical-price-full/{company}?serietype=line'
                            f'&apikey={os.environ["APIKEY"]}')
    data = response.json()['historical']
    return data


def get_all_ratio_data(tickers):
    data = []
    for company in tickers:
        response = requests.get(f'https://financialmodelingprep.com/api/v3/ratios-ttm/{company}?apikey={os.environ["APIKEY"]}')
        # ticker_and_response = [company] + response.json()
        # data.append(ticker_and_response)
        response = response.json()
        response[0]['company'] = company
        data.append(response)
    return data


def get_ratio(data, ticker, ratio):
    """returns the financial valuation ratio specified for a company. See possible_ratios.py for a complete list
    'data' expects the output of get_all_ratio_data
    """
    found_ratio = False
    for entry in data:
        if entry[0]['company'] == ticker.upper():
            ratio = entry[0][ratio]
            found_ratio = True
    if not found_ratio:
        ratio = None

    return round(ratio, 1)


def insider_trading(ticker):
    pass

