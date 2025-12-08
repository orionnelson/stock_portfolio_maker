import os
import json
from dotenv import load_dotenv
from google import genai

load_dotenv()

# Initialize the client with API key
client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])

def generate_portfolio_config(port_request: str) -> dict:
    """
    Generate a portfolio configuration JSON from a natural language request using Gemini.
    """

    generation_config = {
        "temperature": 1,
        "top_p": 0.95,
        "top_k": 64,
        "max_output_tokens": 8192,
        "response_mime_type": "text/plain",
    }

    # Prompt for Gemini
    prompt = f"""
    A user wants to build a portfolio with the following API; you must produce the JSON config for the portfolio.

    The user says: '{port_request}' 

    Follow this format strictly:
    - DO NOT PROVIDE EXPLANATION — ONLY RETURN JSON
    - DO NOT CALCULATE NUMBER OF STOCKS
    - Use these fields:
        • total_value
        • num_picks
        • flat_fee
        • filter_pe_ratio
        • pe_ratio_threshold
        • output_filename
        • performance_period
        • top_n_sectors
        • banned_sectors
        • historical
        • currency
        • cache
        • picks (dictionary of stock tickers to allocations)

    Example:
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
        "historical": "5y",
        "_comment": "Period '6m' is invalid, must be one of ['1d', '5d', '1mo', '3mo', '6mo', '1y', '2y', '5y', '10y', 'ytd', 'max']",
        "currency": "CAD",
        "cache": "lazy",
        "picks": {{"TSLA":20,"GOOG":30}}
    }}
    """

    # Call Gemini
    response = client.models.generate_content(
        model="gemini-1.5-flash",  # Use gemini-1.5-pro if you need better quality
        contents=prompt,
        config=generation_config,
    )

    # Clean response (remove markdown)
    cleaned_lines = [
        line for line in response.text.splitlines()
        if "```" not in line
    ]
    response_text = "\n".join(cleaned_lines).strip()
    print(response_text)

    # Parse JSON
    try:
        portfolio_config = json.loads(response_text)
    except json.JSONDecodeError:
        raise ValueError("Failed to parse the response as JSON. Response text: " + response_text)

    # Validate required fields
    required_keys = [
        "total_value", "num_picks", "flat_fee", "filter_pe_ratio", 
        "pe_ratio_threshold", "performance_period", "historical", "currency", "picks"
    ]
    for key in required_keys:
        if key not in portfolio_config:
            raise ValueError(f"Gemini Failed: Missing required key '{key}' in the portfolio configuration.")

    return portfolio_config
