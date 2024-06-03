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

```json
{
    "total_value": 32000,
    "num_picks": 10,
    "flat_fee": 9.99,
    "filter_pe_ratio": true,
    "pe_ratio_threshold": 25,
    "output_filename": "top_stocks_ytd.xlsx",
    "performance_period": "ytd",
    "top_n_sectors": 4,
    "banned_sectors": ["Energy"]
  }