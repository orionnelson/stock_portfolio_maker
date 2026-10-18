# QA evidence — Qwen 3.6 / Airflow upgrade

## Scope and baseline

- Repository: `orionnelson/stock_portfolio_maker` (public).
- Exact release baseline: `main` at `3a4fa9ade7f0b8261d76b5f88f601fd068564be2`.
- Baseline contains merged PR #14, `Fix excluded tickers`.
- Read-only GitHub API retrieval reconstructed the exact SHA-verified signed commit, tree and seven source blobs into a shallow local Git repository because OpenCode's HTTPS Git helper is broken. No synthetic substitute baseline was committed.
- Local integration branch `blue-master` starts at that exact commit; implementation branch `blue-qwen36-airflow` branches from it. No integration/release merges.

## Executed checks

- **69 pytest tests passed** on Python 3.11/Linux; no skips in the executed suite.
- Actual loopback HTTP request/response exercises the OpenAI-compatible adapter and verifies model ID, endpoint and JSON envelope. Redirects are rejected.
- Malformed/truncated/oversized responses, bad configurations, unsafe filenames, duplicate LLM JSON fields, provider timeout without cloud fallback and narrow tool arguments are covered.
- 500 deterministic randomized allocation cases assert nonnegative integer shares and no overspending. Explicit weight/cash/epsilon/fee regressions are covered.
- Actual legacy selection/ranking runs with synthetic upstream fundamentals.
- Actual openpyxl workbook export and JSON summary generation run with mocked market-data boundaries.
- Actual Prophet models fit synthetic historical prices and produce a workbook containing the forecast sheet/image. No external price data or performance claims.
- Manual CLI `--config config.json --validate-only` succeeded without LLM/market calls.
- Project Python modules, DAG and tests compile successfully.
- `pip check` and `git diff --check` pass. `requirements-tested.lock` records the tested Linux environment, including pytest; Windows COM is not in that lock.

## Independent review

A separate OpenCode ChatGPT session reviewed the new provider, validator, allocation, CLI and DAG-adapter code. It found three reproducible allocation defects in the initial draft:

1. Remainder redistribution could override explicit weights/allocated cash.
2. Floating-point epsilon above 100% could overspend at large budgets.
3. Fees could remain reserved for positions receiving zero shares.

All three were addressed: per-position percentage caps include actual fees, unused cash remains cash, epsilon-only excess is normalized with Decimal arithmetic, and the final exact budget invariant is checked. Dedicated regression tests pass. This is evidence of review and fixes, not a whole-repository security certification.

## Not verified / required acceptance gates

- No live shared Qwen inference: read-only model catalog confirmed the configured Qwen 3.6 Q4 no-think ID exists and was loaded, but no load/unload/runtime change or real generation was performed.
- No live Gemini request: SDK boundary is mocked; model availability, billing and real authentication remain operator checks.
- No live market-data run: Wikipedia/Yahoo/FX availability and schema drift are not covered by synthetic fixtures.
- DAG adapter behavior is tested with an explicitly stubbed Airflow SDK. **No real Airflow scheduler/DagBag parse or worker execution was performed.** Validate with your Airflow 3.3.2 worker image and official dependency constraints before deploying.
- No Windows COM integration test or cross-platform lock reproduction.
- Existing daily pickle caches are trusted-local-only and remain inherited technical debt. Never read them from untrusted/shared writable storage. Existing forecasts/rankings are not validated investment models.
- No CI execution on GitHub: workflow file is prepared only.
- GitHub CLI is unauthenticated. No branches pushed, no PR created, no deployment or live scheduling. Publishing requires securely provisioned GitHub App installation authentication.
