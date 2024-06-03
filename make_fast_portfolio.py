import yfinance as yf
import pandas as pd
from collections import defaultdict
import requests
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import json

def load_config(file_path):
    with open(file_path, 'r') as file:
        config = json.load(file)
    return config

def get_sp500_tickers():
    url = 'https://en.wikipedia.org/wiki/List_of_S%26P_500_companies'
    response = requests.get(url)
    tables = pd.read_html(response.text)
    sp500_table = tables[0]
    tickers = sp500_table['Symbol'].tolist()
    return tickers, sp500_table

def categorize_by_sector(sp500_table):
    sector_dict = defaultdict(list)
    for _, row in sp500_table.iterrows():
        sector_dict[row['GICS Sector']].append(row['Symbol'])
    return sector_dict

def get_sector_performance(sector_dict, period):
    sector_performance = {}
    
    for sector, tickers in sector_dict.items():
        tickers_str = " ".join(tickers)
        sector_data = yf.download(tickers_str, period=period, group_by='ticker', auto_adjust=True)
        sector_close = sector_data.xs('Close', axis=1, level=1)
        sector_performance[sector] = (sector_close.iloc[-1].sum() / sector_close.iloc[0].sum()) - 1
    
    return pd.Series(sector_performance, name="Performance")

def analyze_fundamentals(stock_ticker):
    stock = yf.Ticker(stock_ticker)
    info = stock.info
    
    eps = info.get("trailingEps")
    revenue = info.get("totalRevenue")
    market_cap = info.get("marketCap")
    pe_ratio = info.get("trailingPE")
    price_to_sales = info.get("priceToSalesTrailing12Months")
    price_to_book = info.get("priceToBook")
    ebitda = info.get("ebitda")
    profit_margins = info.get("profitMargins")
    share_price = info.get("currentPrice")
    
    return {
        "EPS": eps,
        "Revenue": revenue,
        "MarketCap": market_cap,
        "PE_Ratio": pe_ratio,
        "Price_to_Sales": price_to_sales,
        "Price_to_Book": price_to_book,
        "EBITDA": ebitda,
        "Profit_Margins": profit_margins,
        "Share_Price": share_price
    }

def get_exchange_rate():
    url = 'https://api.exchangerate-api.com/v4/latest/USD'
    response = requests.get(url)
    data = response.json()
    return data['rates']['CAD']

def select_top_stocks(config):
    tickers, sp500_table = get_sp500_tickers()
    sector_dict = categorize_by_sector(sp500_table)
    sector_performance = get_sector_performance(sector_dict, config["performance_period"])
    print("Sector Performance (YTD):\n", sector_performance)
    
    top_sectors = sector_performance.nlargest(3).index.tolist()
    print("Top Sectors:", top_sectors)
    
    stock_candidates = []
    
    for sector in top_sectors:
        stocks = sector_dict[sector]
        print(f"Stocks in sector {sector}: {stocks}")
        
        for stock in stocks:
            fundamentals = analyze_fundamentals(stock)
            print(f"Fundamentals for {stock}: {fundamentals}")
            stock_candidates.append({"Ticker": stock, **fundamentals})
    
    df = pd.DataFrame(stock_candidates)
    print("DataFrame before dropping NaNs:\n", df)
    
    df = df.dropna()
    print("DataFrame after dropping NaNs:\n", df)
    
    numeric_columns = ["EPS", "Revenue", "MarketCap", "PE_Ratio", "Price_to_Sales", "Price_to_Book", "EBITDA", "Profit_Margins", "Share_Price"]
    for col in numeric_columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    df = df.dropna()
    print("DataFrame after converting to numeric and dropping NaNs:\n", df)
    
    df = df[(df["EPS"] > 0) & (df["Revenue"] > 0) & (df["MarketCap"] > 0) & (df["PE_Ratio"] > 0)]
    
    if config["filter_pe_ratio"]:
        df = df[df["PE_Ratio"] < config["pe_ratio_threshold"]]
    
    print("DataFrame after filtering positive values and PE Ratio:\n", df)
    
    # Rank and normalize the values
    df["EPS_Rank"] = df["EPS"].rank(ascending=False)
    df["Revenue_Rank"] = df["Revenue"].rank(ascending=False)
    df["MarketCap_Rank"] = df["MarketCap"].rank(ascending=False)
    
    # Normalize the ranks
    df["EPS_Score"] = df["EPS_Rank"] / df["EPS_Rank"].max()
    df["Revenue_Score"] = df["Revenue_Rank"] / df["Revenue_Rank"].max()
    df["MarketCap_Score"] = df["MarketCap_Rank"] / df["MarketCap_Rank"].max()
    
    # Calculate the final score
    df["Score"] = df["EPS_Score"] + df["Revenue_Score"] + df["MarketCap_Score"]
    
    df["Score"] = pd.to_numeric(df["Score"], errors='coerce')
    
    df = df.dropna(subset=["Score"])
    print("DataFrame after calculating Score and dropping NaNs:\n", df)
    
    top_stocks = df.nsmallest(config["num_picks"], "Score")
    
    return top_stocks, sector_performance

"""def build_portfolio(top_stocks, config):
    exchange_rate = get_exchange_rate()
    flat_fee = config["flat_fee"]
    top_stocks['Share_Price_CAD'] = top_stocks['Share_Price'] * exchange_rate
    top_stocks['Num_Shares'] = (config["total_value"] / (top_stocks['Share_Price_CAD'] + flat_fee)).apply(lambda x: int(x / len(top_stocks)))
    top_stocks['Investment'] = (top_stocks['Num_Shares'] * top_stocks['Share_Price_CAD']) + flat_fee
    return top_stocks"""

def build_portfolio(top_stocks, config):
    exchange_rate = get_exchange_rate()
    flat_fee = config["flat_fee"]
    total_value = config["total_value"]

    # Convert share prices to CAD
    top_stocks['Share_Price_CAD'] = top_stocks['Share_Price'] * exchange_rate

    # Calculate the initial evenly distributed investment amount
    initial_investment_per_stock = total_value / len(top_stocks)

    # Initialize the number of shares to buy for each stock
    top_stocks['Num_Shares'] = (initial_investment_per_stock / (top_stocks['Share_Price_CAD'] + flat_fee)).apply(lambda x: int(x))

    # Calculate the initial investment for each stock
    top_stocks['Investment'] = (top_stocks['Num_Shares'] * top_stocks['Share_Price_CAD']) + flat_fee

    # Calculate the remaining value to invest
    total_invested = top_stocks['Investment'].sum()
    remaining_value = total_value - total_invested

    # Adjust for the remainder to distribute any leftover funds
    while remaining_value >= top_stocks['Share_Price_CAD'].min():
        for index, row in top_stocks.iterrows():
            additional_investment = row['Share_Price_CAD'] 
            if remaining_value >= additional_investment:
                top_stocks.at[index, 'Num_Shares'] += 1
                top_stocks.at[index, 'Investment'] += additional_investment
                remaining_value -= additional_investment

    return top_stocks


def add_to_excel(top_stocks, sector_performance, filename):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Top Stocks"
    
    for r in dataframe_to_rows(top_stocks, index=False, header=True):
        ws1.append(r)
    
    ws2 = wb.create_sheet(title="Sector Performance")
    sector_df = pd.DataFrame(sector_performance, columns=['Performance'])
    for r in dataframe_to_rows(sector_df.reset_index(), index=False, header=True):
        ws2.append(r)
    
    wb.save(filename)

if __name__ == "__main__":
    config = load_config('config.json')
    top_stocks, sector_performance = select_top_stocks(config)
    print("Top Stocks:\n", top_stocks)
    
    top_stocks = build_portfolio(top_stocks, config)
    print("Portfolio:\n", top_stocks)
    
    add_to_excel(top_stocks, sector_performance, config["output_filename"])
    print(f"Data saved to '{config['output_filename']}'")
