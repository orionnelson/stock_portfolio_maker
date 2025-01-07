# Stock Portfolio Builder

This project is designed to help build a stock portfolio by selecting top-performing S&P 500 stocks based on various fundamental metrics, and distributing a given total investment value among the selected stocks.

## Features

- Fetches S&P 500 tickers from Wikipedia.
- Categorizes stocks by sector.
- Analyzes fundamental metrics such as EPS, Revenue, Market Cap, PE Ratio, etc.
- Ranks and scores stocks based on their fundamental metrics.
- Selects top stocks based on the highest scores.
- Allocates shares of the selected stocks based on a given total investment value.
- Ensures efficient distribution of remaining funds.

## Configuration

The configuration parameters are stored in a `config.json` file. Below is an example of the configuration file:
This configeration plugged in by gemini expermental in the repos Experimental Mode

- Get an API key from `https://aistudio.google.com/app/prompts/new_chat?model=gemini-exp-1206` And add it to the `.env` as `GEMINI_API_KEY=YOUR_KEY`
- Then `python -m venv trader && trader\Scripts\activate && pip install -r requirements.txt`
- To Run `python make_fast_portfolio.py`
- `Input Example`:  "Id like a portfolio consisting of 500000 swedish kronery that is 15% RYCEY 10% Microsoft 20% Google and 10 stocks total based on 3 month performance period and 5y history pe ratio threshold 15"
- `Returns`: Excel File

```json
{
        "total_value": 38000,
        "num_picks": 10,
        "flat_fee": 9.99,
        "filter_pe_ratio": true,
        "pe_ratio_threshold": 25,
        "output_filename": "top_stocks_1y.xlsx", 
        "performance_period": "1y",
        "top_n_sectors": 4,
        "banned_sectors": ["Energy","Financials"], 
        "historical" : "5y",
        "_comment": "Period '6m' is invalid, must be one of ['1d', '5d', '1mo', '3mo', '6mo', '1y', '2y', '5y', '10y', 'ytd', 'max'",
        "currency": "CAD",
        "cache": "lazy",
        "picks":{{"TSLA":20,"GOOG":30}}  
    }
```

- `total_value`: Amount You Are Willing to Spend on your Portfolio total should be under this amount
- `num_picks` : Number of automated picks your portfolio will be divided among
- `flat_fee` : Flat price of a trade which charged by broker.
- `filter_pe_ratio`: Option to filter by the PE threshold avoid overvalued overly speculative stocks
- `pe_ratio_threshold` : Maximum PE Ratio  Threshold limit.
- `performance_period` : Run Initial Calculations for fundamentals over this period 
- `top_n_sectors` : Take n number best performing sectors from the S&P500 and use them for stock picks.
- `banned_sectors`: Avoid a S&P500 Sector
- `historical`: Time Period Used for Prophet Prediction. 
- `currency`: Currency for the Portfolio. CAD USD EUR etc

