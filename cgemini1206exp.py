import os
import google.generativeai as genai 
from dotenv import load_dotenv
import json

load_dotenv()
genai.configure(api_key=os.environ["GEMINI_API_KEY"])
# Create the model https://aistudio.google.com/app/prompts/new_chat?model=gemini-exp-1206
# Add the Key to .env 

def generate_portfolio_config(port_request: str) -> dict:


    generation_config = {
    "temperature": 1,
    "top_p": 0.95,
    "top_k": 64,
    "max_output_tokens": 8192,
    "response_mime_type": "text/plain",
    } 

    model = genai.GenerativeModel(
    model_name="gemini-exp-1206",
    generation_config=generation_config,
    ) 

    chat_session = model.start_chat(
    history=[
    ]
    ) 


    response = chat_session.send_message(f"""A user wants to build a portfolio with the following api however you must come up with the json config for the portfolio they want.


    The user says '{port_request}' come up with the json schema for this portfolio based on below info and json example 

    • total_value: Amount You Are Willing to Spend on your Portfolio total should be under this amount
    • num_picks : Number of automated picks your portfolio will be divided among
    • flat_fee : Flat price of a trade which charged by broker.
    • filter_pe_ratio: Option to filter by the PE threshold avoid overvalued overly speculative stocks
    • pe_ratio_threshold : Maximum PE Ratio Threshold limit.
    • performance_period : Run Initial Calculations for fundamentals over this period
    • top_n_sectors : Take n number best performing sectors from the S&P500 and use them for stock picks.
    • banned_sectors: Avoid a S&P500 Sector
    • historical: Time Period Used for Prophet Prediction.
    • currency: Currency for the Portfolio. CAD USD EUR etc


    {{
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
    }}
    """) 


    cleaned_lines = [
        line for line in response.text.splitlines()
        if "```" not in line
    ]
    response_text = "\n".join(cleaned_lines)

    # Parse the response text as JSON
    print(response_text)

    try:
        portfolio_config = json.loads(response_text)
    except json.JSONDecodeError:
        raise ValueError("Failed to parse the response as JSON. Response text: " + response_text)

    # Validate the structure of the JSON
    required_keys = [
        "total_value", "num_picks", "flat_fee", "filter_pe_ratio", 
        "pe_ratio_threshold", "performance_period", "top_n_sectors", 
        "banned_sectors", "historical", "currency", "picks"
    ]

    for key in required_keys:
        if key not in portfolio_config:
            raise ValueError(f" Gemini Failed: Missing required key '{key}' in the portfolio configuration.")

    return portfolio_config



