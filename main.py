from financial_modeling_prep_api import portfolio, CASH, get_closing_price_and_volume, get_historical_closing_prices, \
    get_portfolio_value, get_all_ratio_data, get_ratio


def main():
    price_volume = get_closing_price_and_volume(portfolio.keys())
    portfolio_value = get_portfolio_value(price_volume)
    # ratio_data = get_all_ratio_data(portfolio.keys())
    print(price_volume)
    print(portfolio_value)
    print(float(portfolio_value) - float(CASH))


if __name__ == "__main__":
    main()

