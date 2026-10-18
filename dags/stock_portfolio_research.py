"""Manual-only Airflow 2/3 port. Install repo dependencies in the worker image."""
from datetime import datetime, timezone, timedelta
import hashlib
import json
import os
from pathlib import Path

try:
    from airflow.sdk import dag, task, get_current_context
except ImportError:
    from airflow.decorators import dag, task
    from airflow.operators.python import get_current_context


def execute_run(conf, run_id):
    """Use only operator-controlled filesystem roots and provider endpoints."""
    from portfolio_config import validate_config
    from portfolio_cli import run_portfolio
    if not isinstance(conf, dict) or set(conf) - {'config', 'request', 'provider', 'forecasts'}:
        raise ValueError('Unsupported DAG configuration')
    if 'request' in conf and 'config' in conf:
        raise ValueError('Specify either config or request')
    if type(conf.get('forecasts', True)) is not bool:
        raise ValueError('forecasts must be boolean')
    if 'request' in conf:
        from portfolio_llm import generate_portfolio_config
        config = generate_portfolio_config(conf['request'], provider=conf.get('provider', 'qwen'))
    elif 'config' in conf:
        config = validate_config(conf['config'])
    else:
        config_path = Path(os.environ.get('PORTFOLIO_CONFIG_PATH', '/opt/stock-portfolio/config.json'))
        config = validate_config(json.loads(config_path.read_text()))
    root = Path(os.environ.get('PORTFOLIO_OUTPUT_ROOT', '/opt/airflow/data/portfolios'))
    # Unique deterministic paths for separate manual runs; no user-controlled path fragments.
    output_dir = root / hashlib.sha256(run_id.encode()).hexdigest()[:24]
    return run_portfolio(config, output_dir=output_dir, forecasts=conf.get('forecasts', True), plots=False)


@dag(dag_id='stock_portfolio_research', schedule=None,
     start_date=datetime(2026, 1, 1, tzinfo=timezone.utc), catchup=False,
     max_active_runs=1, tags=['research', 'portfolio'], default_args={'retries': 0})
def stock_portfolio_research():
    @task(execution_timeout=timedelta(minutes=45))
    def build():
        context = get_current_context()
        return execute_run(context['dag_run'].conf or {}, context['run_id'])
    build()

stock_portfolio_research()
