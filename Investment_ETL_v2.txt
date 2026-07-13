import pandas as pd
import os

excel_path = "/lakehouse/default/Files/portfolio.xlsx"


excel_path = "/lakehouse/default/Files/portfolio.xlsx"
print(excel_path)
print(os.listdir("/lakehouse/default/Files"))


xls = pd.ExcelFile(excel_path)

print(xls.sheet_names)


portfolio_pdf = pd.read_excel(
    excel_path,
    sheet_name="⚡ PowerBI_Ready"
)

deposits_pdf = pd.read_excel(
    excel_path,
    sheet_name="💵 Deposits"
)

dividends_pdf = pd.read_excel(
    excel_path,
    sheet_name="🎁 Dividends"
)

monthly_pdf = pd.read_excel(
    excel_path,
    sheet_name="📅 Monthly"
)

ticker_pdf = pd.read_excel(
    excel_path,
    sheet_name="🏷️ By Ticker"
)

all_trades_pdf = pd.read_excel(
    excel_path,
    sheet_name="📋 All Trades"
)


display(portfolio_pdf.head())

display(deposits_pdf.head())

display(dividends_pdf.head())

display(monthly_pdf.head())

display(ticker_pdf.head())

display(all_trades_pdf.head())


def clean_columns(df):
    df.columns = (
        df.columns
        .str.strip()
        .str.replace(" ", "_", regex=False)
        .str.replace("(", "", regex=False)
        .str.replace(")", "", regex=False)
        .str.replace("$", "USD", regex=False)
        .str.replace("%", "Pct", regex=False)
        .str.replace("-", "_", regex=False)
        .str.replace("/", "_", regex=False)
    )
    return df


portfolio_pdf = clean_columns(portfolio_pdf)
deposits_pdf = clean_columns(deposits_pdf)
dividends_pdf = clean_columns(dividends_pdf)
monthly_pdf = clean_columns(monthly_pdf)
ticker_pdf = clean_columns(ticker_pdf)
all_trades_pdf = clean_columns(all_trades_pdf)


def save_table(pdf, table_name):
    spark.createDataFrame(pdf) \
        .write \
        .format("delta") \
        .mode("overwrite") \
        .option("mergeSchema", "true") \
        .saveAsTable(table_name)


save_table(portfolio_pdf, "portfolio_powerbi")
save_table(deposits_pdf, "deposits")
save_table(dividends_pdf, "dividends")
save_table(monthly_pdf, "monthlysummary")
save_table(ticker_pdf, "tickersummary")
save_table(all_trades_pdf, "all_trades")


spark.sql("SHOW TABLES").show(truncate=False)


import pandas as pd
from pyspark.sql import functions as F

# Load transactions
df = spark.table("portfolio_powerbi").toPandas()

# Clean columns
df["DateTime"] = pd.to_datetime(df["DateTime"], errors="coerce")
df["Side"] = df["Side"].str.upper().str.strip()
df["Ticker"] = df["Ticker"].str.upper().str.strip()
df["Shares"] = pd.to_numeric(df["Shares"], errors="coerce").fillna(0)

df = df.sort_values(["Ticker", "DateTime"])

# Calculate current shares only
positions = []

for ticker, group in df.groupby("Ticker"):
    current_shares = 0.0

    for _, row in group.iterrows():
        side = row["Side"]
        shares = float(row["Shares"])

        if side == "BUY":
            current_shares += shares

        elif side == "SELL":
            current_shares -= shares

    positions.append({
        "Ticker": ticker,
        "Current_Shares": current_shares
    })

positions_df = pd.DataFrame(positions)

# Remove tiny leftovers
positions_df = positions_df[positions_df["Current_Shares"] > 0.001]

# Load Trading212-style average buy from ticker summary
ticker_summary = spark.table("tickersummary").toPandas()
ticker_summary["Ticker"] = ticker_summary["Ticker"].str.upper().str.strip()
ticker_summary["Avg_Buy_USD"] = pd.to_numeric(
    ticker_summary["Avg_Buy_USD"],
    errors="coerce"
).fillna(0)

positions_df = positions_df.merge(
    ticker_summary[["Ticker", "Avg_Buy_USD"]],
    on="Ticker",
    how="left"
)

positions_df["Average_Cost"] = positions_df["Avg_Buy_USD"]
positions_df["Remaining_Cost_Basis"] = (
    positions_df["Current_Shares"] * positions_df["Average_Cost"]
)

positions_df = positions_df.drop(columns=["Avg_Buy_USD"])

# Load latest stock prices
prices = spark.table("stock_prices").toPandas()
prices["Ticker"] = prices["Ticker"].str.upper().str.strip()

prices["Date"] = pd.to_datetime(prices["Date"], errors="coerce")
prices = prices.sort_values("Date").groupby("Ticker").tail(1)

prices = prices[["Ticker", "Current_Price", "Close", "Previous_Close", "Date"]]

# Merge positions with prices
final_df = positions_df.merge(prices, on="Ticker", how="left")

final_df["Market_Value"] = final_df["Current_Shares"] * final_df["Current_Price"]
final_df["Unrealized_Profit"] = (
    final_df["Market_Value"] - final_df["Remaining_Cost_Basis"]
)
final_df["Unrealized_Return_%"] = (
    final_df["Unrealized_Profit"] / final_df["Remaining_Cost_Basis"]
)

final_df = final_df.replace([float("inf"), float("-inf")], 0).fillna(0)

display(final_df)

spark_final = spark.createDataFrame(final_df)

spark_final.write.mode("overwrite") \
    .format("delta") \
    .option("overwriteSchema", "true") \
    .saveAsTable("portfolio_positions")

print("portfolio_positions table created successfully")

display(spark_final)