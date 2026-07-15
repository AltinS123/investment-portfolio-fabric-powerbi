pip install yfinance

# Welcome to your new notebook
# Type here in the cell editor to add code!

import yfinance as yf
import pandas as pd
from datetime import datetime

import yfinance as yf
import pandas as pd
from datetime import datetime


stocks = [
 "AMZN",
 "AMD",
 "NVDA",
 "GOOGL",
 "IBM",
 "META",
 "AAPL",
 "SOFI",
 "DVLT",
 "OPEN",
 "IBO",
 "TSLA",
 "GDHG",
 "OPENZ",
 "OPENW",
 "HCTI",
 "OPENL",
 "NKE",
 "SHEL",
 "NFLX",
 "MCD",
 "RGTI",
 "PLTR",
 "NOW",
 "NOK",
 "SPCX"
]


stock_data = []


for stock in stocks:

    ticker = yf.Ticker(stock)

    # Get latest historical data
    hist = ticker.history(period="5d")

    close_price = hist["Close"].iloc[-1]


    stock_data.append({

        "Ticker": stock,

        "Current_Price": float(ticker.fast_info["lastPrice"]),

        "Close": float(close_price),

        "Previous_Close": float(ticker.fast_info["previousClose"]),

        "Day_High": float(ticker.fast_info["dayHigh"]),

        "Day_Low": float(ticker.fast_info["dayLow"]),

        "Date": datetime.now()
    })


stock_pdf = pd.DataFrame(stock_data)


display(stock_pdf)


stock_spark = spark.createDataFrame(stock_pdf)

stock_spark.write \
    .format("delta") \
    .mode("overwrite") \
    .option("mergeSchema", "true") \
    .saveAsTable("Stock_Prices")



spark.sql("SHOW TABLES").show()




