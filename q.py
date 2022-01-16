import requests
import nasdaqdatalink
import matplotlib.pyplot as plt
import quandl

# api_key = 'L3Vz3zt_BzxxhbCMdDsi'
data = quandl.get('WIKI/GOOGL')
prices = data['Adj. Close']
print(prices)
plt.plot(prices)
plt.show()