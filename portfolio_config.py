"""Validate untrusted manual/LLM configuration before market or file operations."""
import math
import re
from decimal import Decimal

PERIODS = {'1d', '5d', '1mo', '3mo', '6mo', '1y', '2y', '5y', '10y', 'ytd', 'max'}
DEFAULTS = {'flat_fee': 9.99, 'filter_pe_ratio': True, 'pe_ratio_threshold': 25,
            'performance_period': '1y', 'historical': '5y', 'currency': 'CAD',
            'cache': 'lazy', 'top_n_sectors': 4, 'banned_sectors': [],
            'picks': {}, 'output_filename': 'portfolio.xlsx'}

def validate_config(value):
    if not isinstance(value, dict):
        raise ValueError('Configuration must be a JSON object')
    allowed = set(DEFAULTS) | {'total_value', 'num_picks', '_comment'}
    if set(value) - allowed:
        raise ValueError('Unknown configuration fields')
    config = dict(DEFAULTS, **value)
    for key in ('total_value', 'flat_fee', 'pe_ratio_threshold'):
        x = config.get(key)
        if isinstance(x, bool) or not isinstance(x, (int, float)) or not math.isfinite(x):
            raise ValueError(f'{key} must be a finite number')
        if x < 0 or (key != 'flat_fee' and x == 0):
            raise ValueError(f'{key} outside valid range')
    for key, maximum in [('num_picks', 100), ('top_n_sectors', 11)]:
        x = config.get(key)
        if type(x) is not int or not 1 <= x <= maximum:
            raise ValueError(f'{key} must be an integer between 1 and {maximum}')
    if config['flat_fee'] >= config['total_value']:
        raise ValueError('Budget must exceed the fee for one funded position')
    if type(config['filter_pe_ratio']) is not bool:
        raise ValueError('filter_pe_ratio must be boolean')
    for key in ('performance_period', 'historical'):
        if not isinstance(config[key], str) or config[key] not in PERIODS:
            raise ValueError(f'Invalid {key}')
    if config['cache'] not in ('lazy', 'refresh'):
        raise ValueError('cache must be lazy or refresh')
    if not isinstance(config['currency'], str) or not re.fullmatch(r'[A-Z]{3}', config['currency']):
        raise ValueError('currency must be a three-letter uppercase code')
    name = config['output_filename']
    if not isinstance(name, str) or not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_.-]{0,119}\.xlsx', name):
        raise ValueError('output_filename must be a simple .xlsx basename')
    sectors = config['banned_sectors']
    if not isinstance(sectors, list) or any(not isinstance(s, str) or not s.strip() for s in sectors):
        raise ValueError('banned_sectors must be a list of nonempty strings')
    picks = config['picks']
    if isinstance(picks, list):
        if not picks or len(picks) != len(set(str(p) for p in picks)):
            raise ValueError('Ticker list must be nonempty and unique')
        picks = {p: 100 / len(picks) for p in picks if isinstance(p, str)}
        if len(picks) != len(config['picks']):
            raise ValueError('Invalid ticker list')
    if not isinstance(picks, dict) or len(picks) > config['num_picks']:
        raise ValueError('picks must fit within num_picks')
    for ticker, weight in picks.items():
        if not isinstance(ticker, str) or not re.fullmatch(r'[A-Z][A-Z0-9.^=-]{0,14}', ticker):
            raise ValueError('Invalid ticker')
        if type(weight) not in (int, float) or not math.isfinite(weight) or not 0 < weight <= 100:
            raise ValueError('Allocations must be finite percentages in (0,100]')
    total_weight = sum(Decimal(str(w)) for w in picks.values())
    if total_weight > Decimal('100.00000001'):
        raise ValueError('Allocations exceed 100 percent')
    config['picks'] = picks
    return config
