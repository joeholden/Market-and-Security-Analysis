from math import log, sqrt, pi, exp
from scipy.stats import norm
from datetime import datetime, date
import numpy as np
import pandas as pd
from pandas import DataFrame
import matplotlib.pyplot as plt

# Black-Scholes pricing is for European options. American options can be priced using the Bjerksund-Stensland Model


def d1(s, k, t, r, sigma):
    return (log(s / k) + (r + sigma ** 2 / 2.) * t) / (sigma * sqrt(t))


def d2(s, k, t, r, sigma):
    return d1(s, k, t, r, sigma) - sigma * sqrt(t)


def call(s, k, t, r, sigma):
    """Returns price of a call option from Black-Scholes"""
    return s * norm.cdf(d1(s, k, t, r, sigma)) - k * exp(-r * t) * norm.cdf(d2(s, k, t, r, sigma))


def put(S, K, T, r, sigma):
    """Returns price of a put option from Black-Scholes"""
    return K * exp(-r * T) - S + call(S, K, T, r, sigma)