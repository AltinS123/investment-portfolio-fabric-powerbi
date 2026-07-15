Unrealized Return % =
DIVIDE(
    [Unrealized Profit],
    [Cost Basis]
)


Unrealized Return % =
DIVIDE(
    [Unrealized Profit],
    [Cost Basis]
)


Dividends =
SUM(dividends[Div_Share_USD])


Total Return =
[Unrealized Profit]
    + [Realized Profit]
    + [Dividends]



Total Return % =
DIVIDE(
    [Total Return],
    [Net Cash Invested]
)



Last Refresh =
"Last refresh: "
    & FORMAT(
        MAX(stock_prices[Date]),
        "dd MMM yyyy HH:mm"
    )



Current Shares =
SUM(portfolio_positions[Current_Shares])


Net Cash Invested =
SUM(deposits[Amount_USD])


Realized Profit by Ticker =
SUM(tickersummary[Realized_P&L_USD])
