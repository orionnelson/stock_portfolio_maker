"""Interactive compatibility entry point and noninteractive CLI/Airflow pipeline."""
import argparse
import json
from pathlib import Path
from portfolio_config import validate_config


def run_portfolio(config, *, output_dir='.', forecasts=True, plots=False):
    config = validate_config(config)
    # Select the headless backend before importing the legacy plotting module.
    if not plots:
        import matplotlib
        matplotlib.use('Agg')
    import make_fast_portfolio as engine
    output = Path(output_dir).resolve()
    output.mkdir(parents=True, exist_ok=True)
    destination = output / config['output_filename']
    picks = config['picks']
    selected, sectors = (engine.select_top_stocks(config, stock_picks=picks) if picks
                         else engine.select_top_stocks(config))
    if selected.empty:
        raise ValueError('No stocks matched the selected criteria')
    portfolio = engine.build_portfolio(selected.copy(), config)
    if plots:
        engine.plot_sector_open_close(config['performance_period'], mode=config['cache'])
    if forecasts:
        engine.save_forecast_to_excel_predict(portfolio, sectors, str(destination), config)
    else:
        engine.add_to_excel(portfolio, sectors, str(destination))
    invested = float(portfolio['Investment'].sum())
    summary = {'output_file': str(destination), 'positions': len(portfolio),
               'budget': config['total_value'], 'invested_including_fees': invested,
               'cash_remaining': round(config['total_value'] - invested, 8),
               'currency': config['currency'], 'forecasts': forecasts,
               'notice': 'Research only; not investment advice. No orders are placed.'}
    destination.with_suffix('.summary.json').write_text(json.dumps(summary, indent=2) + '\n')
    return summary


def main(argv=None):
    parser = argparse.ArgumentParser(description='Build an S&P 500 research portfolio; never places orders.')
    parser.add_argument('--config', help='Manual JSON configuration; bypasses all LLMs')
    parser.add_argument('--request', help='Natural-language portfolio request')
    parser.add_argument('--provider', choices=['gemini', 'qwen'], default='gemini')
    parser.add_argument('--model', help='Explicit model ID on the selected provider')
    parser.add_argument('--base-url', help='Trusted Qwen OpenAI-compatible endpoint')
    parser.add_argument('--output-dir', default='.')
    parser.add_argument('--validate-only', action='store_true', help='Validate/generate configuration without market calls')
    parser.add_argument('--no-forecasts', action='store_true')
    parser.add_argument('--plots', action='store_true', help='Enable interactive sector charts (off for headless runs)')
    args = parser.parse_args(argv)
    if args.config and args.request:
        parser.error('--config and --request are mutually exclusive')
    if args.config:
        config = validate_config(json.loads(Path(args.config).read_text()))
    else:
        from portfolio_llm import generate_portfolio_config
        request = args.request or input('What portfolio would you like?: ')
        config = generate_portfolio_config(request, provider=args.provider, model=args.model, base_url=args.base_url)
    if args.validate_only:
        print(json.dumps(config, indent=2))
        return config
    summary = run_portfolio(config, output_dir=args.output_dir,
                            forecasts=not args.no_forecasts, plots=args.plots)
    print(json.dumps(summary, indent=2))
    return summary

if __name__ == '__main__':
    main()
