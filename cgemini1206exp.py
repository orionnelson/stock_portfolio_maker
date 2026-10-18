"""Backward-compatible Gemini helper; cloud credentials are loaded only on use."""
from portfolio_llm import generate_portfolio_config as _generate


def generate_portfolio_config(port_request: str) -> dict:
    return _generate(port_request, provider="gemini")
