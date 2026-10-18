# Stock Portfolio Builder

Build an S&P 500 **research** portfolio using sector performance, fundamental ranking (including Quick Ratio), optional selected tickers/percentages, whole-share allocation, Excel output and Prophet forecasts. No brokerage integration or trade execution. Forecasts are estimates, not investment advice.

## Install

Python 3.11 is the tested runtime. Use a project-local virtual environment:

```sh
python -m venv .venv
. .venv/bin/activate
python -m pip install -r requirements.txt
```

On Windows activate `.venv\Scripts\activate`; `pywin32` is installed only on Windows. Linux/headless runs do not require Excel COM. Gemini uses the `google-genai` SDK matching `from google import genai`, not the older `google-generativeai` package.

## Manual, local Qwen, or Gemini

```sh
# Offline configuration validation: no LLM, market requests or Excel writes.
python portfolio_cli.py --config config.json --validate-only

# Noninteractive manual mode; skip expensive forecasts if desired.
python portfolio_cli.py --config config.json --output-dir ./reports --no-forecasts

# Optional local Qwen 3.6 (OpenAI-compatible smol friend endpoint).
export SMOL_LLM_BASE_URL=http://127.0.0.1:8000/v1
export SMOL_LLM_MODEL=Qwen3.6-35B-A3B-UD-Q4_K_M-no-think
python portfolio_cli.py --provider qwen --request \
  'CAD 38000, 10 stocks, 30% RYCEY, 1y performance, 5y history' --validate-only

# Cloud Gemini, only when selected. Keep the key out of source control.
export GEMINI_API_KEY=your-key
python portfolio_cli.py --provider gemini --request 'CAD 38000 across ten stocks'
```

Copy `.env.example` to `.env` if desired; existing environment variables take precedence. Set the Qwen endpoint to a trusted address reachable from **your** process/container. No public cloud fallback occurs if Qwen fails. The application does **not** load/unload shared models, alter model settings or automatically retry requests. Provision the chosen model separately. Optional `SMOL_LLM_API_KEY` is supported; use HTTPS when transporting credentials outside a trusted local network. Redirects are not followed.

`python make_fast_portfolio.py` remains the legacy interactive Gemini entry point; it now accepts the same CLI flags. `cgemini1206exp.generate_portfolio_config(request)` remains callable, with lazy credential/SDK initialization. `GEMINI_MODEL` defaults to `gemini-2.5-flash` instead of the retired 1.5 model; choose an account-supported model explicitly when needed.

### Use Qwen as an application-owned LLM tool

```python
from portfolio_llm import TOOL_SCHEMA, run_config_tool

config = run_config_tool(
    {"request": "CAD 38000, 10 stocks, excluding Energy"},
    provider="qwen",
    base_url="http://127.0.0.1:8000/v1",
)
```

`TOOL_SCHEMA` is an OpenAI-style function description for an orchestrating LLM. The host explicitly dispatches tool calls via `run_config_tool`; the repository does not install an autonomous agent loop. Model output cannot select endpoints, files outside the output directory, providers or arbitrary executable tools.

## Configuration and intentional safety changes

`config.json` remains supported. Required fields: positive numeric `total_value`, integer `num_picks` (1–100). Defaults cover trading fee, PE filter, periods, currency, sectors, output name and cache. `picks` accepts a ticker-to-percentage object or a unique ticker list; the latter is converted to equal percentages. Empty `{}` means automatic stock selection. Supported periods: `1d`, `5d`, `1mo`, `3mo`, `6mo`, `1y`, `2y`, `5y`, `10y`, `ytd`, `max`.

- Invalid/non-finite numbers, duplicate JSON fields in LLM responses, unknown fields, allocation totals above 100% (apart from normalized floating epsilon), invalid ticker names and unsafe output paths are rejected before market operations.
- `output_filename` is a simple `.xlsx` basename; set the directory using `--output-dir`, not model output.
- Each position's percentage is now a **hard cap including its one flat trading fee**. Only funded positions pay fees. Zero-weight positions are never funded. Rounding leftovers and unallocated percentages remain cash instead of silently overriding requested weights. This deliberately replaces the old over-budget/remainder redistribution behavior.
- Existing upstream dynamic selection fills missing explicit percentages where enough candidates qualify. Fewer qualifying candidates can leave cash. Selection/ranking logic is otherwise retained.
- Headless operation is the default; `--plots` enables the original interactive sector chart. Forecasts remain enabled unless `--no-forecasts` is supplied.
- Each workbook gains a sibling `.summary.json` reporting budget, spend including fees and cash remaining. It does not contain prompts, API keys or model responses.

## Airflow DAG port

`dags/stock_portfolio_research.py` reuses the **same** validator, provider adapter, selection, allocation and export pipeline. It supports the Airflow 3 SDK with Airflow 2 import fallback. It is manual-only (`schedule=None`, `catchup=False`, one active run, zero automatic retries). Nothing in this repository deploys or triggers it.

Operator setup:

1. Install this repo and its Python dependencies into the Airflow **worker** runtime. Put the repo root on `PYTHONPATH`; install Airflow separately using its official version constraints.
2. Copy/register only the DAG file in your DAG bundle. Avoid running market/network work at DAG parse time.
3. Set trusted worker environment variables `PORTFOLIO_CONFIG_PATH`, `PORTFOLIO_OUTPUT_ROOT`, and optional `SMOL_LLM_*`/`GEMINI_*` settings. Output storage must be writable; the Python worker needs market-data egress for actual builds.
4. Trigger manually with either `{"config": {...}, "forecasts": false}` or `{"request": "CAD 38000, ten stocks", "provider": "qwen"}`. An empty trigger uses the operator-managed JSON config file.
5. Retrieve workbooks/summary files from the configured output storage. Only the small summary crosses XCom. Each run writes under a hashed run-ID subdirectory; no paths/endpoints can be injected through trigger configuration.

The task has a 45-minute execution timeout. Re-running the same run ID reuses its output location. Use dedicated trusted worker storage/CWD: inherited legacy daily pickle caches are not safe to share with untrusted writers. Do not mount uploaded/untrusted pickle files. No task places orders, emails reports or changes production schedules.

## Tests and verification boundaries

```sh
python -m pip install pytest
python -m pytest -q
python -m compileall -q .
```

Tests include real loopback HTTP provider requests, malformed/redirected responses, deterministic budget invariants, manual CLI isolation, synthetic market selection, genuine openpyxl exports and genuine Prophet fits on synthetic price history. The DAG adapter tests use an explicitly stubbed Airflow SDK; they are **not** a real scheduler/DagBag acceptance test. Gemini SDK calls are mocked. No live market, cloud Gemini or shared Qwen inference is claimed by these tests. See `QA_RESULTS.md` for the checked build and outstanding acceptance gates.

## Branch and release policy

Autonomous work: release `main` → `blue-master` → `blue-<feature>` → PR into `blue-master`. **No automatic merges.** A human merges the reviewed feature; only then can `blue-master` → `main` contain it for release review. Developer assistance instead branches from and targets the developer's feature branch. Never retarget directly to `main` to bypass this policy. See `PULL_REQUEST_DRAFT.md` for the review packet modelled on PR #14.
