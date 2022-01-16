def get_mcclellan_index_and_oscillator():
    """
    Ratio-Adjusted Net Advances (RANA): (Advances - Declines)/(Advances + Declines)
    McClellan Oscillator: 19-day EMA of RANA - 39-day EMA of RANA
    The first EMA calculation is a simple average.
    Mcclellan Index is the sum of all previous day's values for the oscillator
    When the Mcclellan Summation Index requires +-4,000 issues to turn, we are oversold/overbought
    """
    data = pd.read_excel('excel files/cleaned.xlsx', engine='openpyxl', names=HEADERS.values())
    mcclellan_summation_index = data['McC Summation Index']
    mcclellan_oscillator = data['McC A-D Osc']
    oscillator_to_zero = data['Osc to zero tomorrow']
    advances = data['NYSE Advancing Issues']
    declines = data['NYSE Declining issues']

    ratio_adjusted_net_advances = (advances-declines) / (advances+declines)

    # The multiplication factor of 3400 is just to get the plot to line up
    # with the published oscillator numbers in the spreadsheet. Empirical
    # More accurate the further from the beginning you go

    data['19-Day EMA'] = ratio_adjusted_net_advances.ewm(span=19).mean() * 3400
    data['39-Day EMA'] = ratio_adjusted_net_advances.ewm(span=39).mean() * 3400
    data['oscillator_j'] = data['19-Day EMA'] - data['39-Day EMA']

    return
    # plt.plot(data['date'], data['oscillator_j'], color='blue')
    # plt.plot(data['date'], mcclellan_oscillator, color='green')
    plt.plot(data['date'], mcclellan_summation_index, color='red')
    plt.plot(data['date'], oscillator_to_zero, color='green')
    plt.show()


get_mcclellan_index_and_oscillator()