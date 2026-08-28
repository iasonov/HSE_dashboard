# CLAUDE.md

Guidance for Claude Code (and other AI assistants) working in this repository.

## Branch note

This repo has no active work on `main` (stale, last touched in March). All
current development happens on **`2026-develop`** — that's the branch this
file describes, and the one to base new work on. There's also a `2026-dev`
branch (further ahead in places, behind in others — not a clean
fast-forward of `2026-develop`) and short-lived feature branches; check with
whoever owns the repo before assuming which one is "current" if `2026-develop`
has moved on since this was written.

## What this project is

A data pipeline that builds a weekly admissions dashboard for HSE Online (Высшая
школа экономики). It aggregates leads/applications/contracts/payments/
enrollments per program, computes conversion metrics, writes an Excel
snapshot, and pushes the result to a shared Google Sheet ("Еженедельный отчет
<year>_общий"). There is no web server or frontend — it's a script run on
demand (weekly, by a human) that produces a spreadsheet.

The pipeline is mid-migration from **file-based** ingestion (Excel/CSV exports
manually downloaded from АСАВ, АИС ПК, Bitrix, the portal) to **Bitrix24 REST
API** ingestion (`bitrix_mode = True` in `src/main.py`). Both paths still
exist and are switched at the top of `if __name__ == '__main__':` in
`src/main.py`.

## Layout

```
src/
  main.py             # entrypoint: run_dashboard() picks bitrix vs. legacy (file) mode, writes
                       # the Excel snapshot, optionally pushes to Google Sheets
  __main__.py          # lets `python -m src` call main.run_dashboard()
  general_pipeline.py  # shared helpers used by both pipelines: categorize_ages, years_ago,
                        # num_years, insert_values, process_by_week, and calculate_dashboard()
                        # (the dispatcher that picks files_pipeline vs. bitrix_pipeline)
  bitrix.py             # low-level read-only Bitrix24 REST client (rate-limited, retries,
                         # batches up to 50 read-only calls) — create_bitrix_client(),
                         # collect_bitrix_item_sources()
  bitrix_contracts.py   # BitrixEntity dataclass: declares, per Bitrix table (crm_deals,
                         # portal_deals, applications, contacts, educational_programs, ...),
                         # which fields to SELECT and what filter to apply. This is the actual
                         # "data contract" for the Bitrix source — read it before changing which
                         # fields the API pipeline pulls.
  bitrix_pipeline.py    # process_from_bitrix(): normalizes the raw Bitrix tables and computes
                         # all dashboard metrics for the current (Bitrix-API) mode
  files_pipeline.py     # process_from_current_files() / process_history_files(): the legacy
                         # file-based pipeline (this used to be process.py), plus
                         # process_exams_dates_file() for a separate "exam dates" export
  contracts_summary.py  # builds a standalone ASAV/АИС ПК contracts-vs-plan summary report
                         # from raw file exports (independent of the main dashboard pipeline)
  update_pipeline.py    # pushes the aggregated DataFrame to Google Sheets via gspread
                         # (update_sheet), plus update_exams_dates_sheet() for the exam-dates sheet
  col_names.py           # single source of truth for every column name (Russian labels copied
                          # from the source systems, plus the internal col_* constants used
                          # throughout the pipelines)
  time_const.py          # a couple of hardcoded admissions-cycle date constants (DATE_01_10_2025,
                          # DATE_01_04_2026) shared by both pipelines
  main_update_exams_dates.py # small standalone script: file_pipeline exam dates -> Google Sheets
  google_test.py (in tests/) # ad-hoc scratch script for testing Google Sheets auth, not part of
                              # the pipeline
  test.ipynb              # exploratory notebook, not automated tests (huge, don't read it whole)
docs/
  data_contracts.md     # data folder layout convention (data/raw, data/processed, data/archive,
                         # data/dashboards) and required columns per file source — read this
                         # alongside bitrix_contracts.py before touching input handling
templates/               # checked-in reference data: program catalog (programs.xlsx), the
                          # dashboard column template (template.xlsx), prior-year exports used
                          # for year-over-year comparisons, entrance-exam dictionaries
tests/                   # pytest/unittest tests, `conftest.py` puts src/ on sys.path.
                          # NOTE (see "Known gaps" below): several of these are currently broken —
                          # they import modules/functions that were renamed or removed during the
                          # Bitrix-API refactor. Don't trust a green `pytest tests/` here without
                          # checking which tests actually collected.
                          # tests/*.php are unrelated Bitrix24 CRest scaffolding, not Python tests.
data/                    # gitignored: raw input files for legacy mode, and the generated
                          # dashboard*.xlsx snapshots
```

Credentials (`service_credentials.json`, `src/my_secrets.py` — holds
`secrets['BITRIX_WEBHOOK_URL']`) are gitignored and never committed. So are
all `data/` contents. See `.gitignore`.

## Running it

```
pip install -r requirements.txt
```
`src/main.py` has a `bitrix_mode` flag at the bottom:
- `True` (current default): pulls straight from the Bitrix24 REST API via
  `bitrix.py`, needs `src/my_secrets.py` with a valid webhook URL.
- `False`: legacy mode — reads exports dropped into `data/` (see
  `docs/data_contracts.md` for exact required columns per source; a missing
  file degrades that metric to 0 rather than erroring).

Either way it writes `data/dashboards/dashboard<timestamp>.xlsx` and, if
`update_dashboard` is `True`, pushes to the Google Sheet using
`service_credentials.json` (via `update_pipeline.update_sheet`).

There's no CI configured. Before committing a change, run the tests that
actually still import cleanly (`pytest tests/ -v` and read the collection
errors) and sanity-check pipeline changes manually — nothing else will catch
a regression.

## Key conventions

- **Column names live only in `col_names.py`.** Every DataFrame column —
  whether it's a literal Russian header copied from a source system
  (`master_col_*`, `bachelor_col_*`, `bitrix_col_*`) or an internal computed
  field (`col_*`) — is a constant defined there and imported via
  `from col_names import *`. Never hardcode a column name string in a
  pipeline module; add or reuse a constant instead.
- **The Bitrix data contract lives in `bitrix_contracts.py`.** Changing what
  fields the API pipeline reads means editing a `BitrixEntity`'s `select`
  there, not adding ad-hoc field access in `bitrix_pipeline.py`.
- **Bitrix client is read-only by construction.** `bitrix.py` only allows
  methods ending in `.get`/`.list`/`.fields` (plus `batch`) —
  `_assert_read_only_method` raises on anything else. Keep it that way; this
  pipeline must never be able to write back to the CRM.
- **Programs are split into "master" (магистратура) and "bachelor"
  (бакалавриат)**, each with their own source columns and dashboard rows,
  concatenated into one dashboard DataFrame at the end of each pipeline.
  `bachelor_dict` in `col_names.py` maps raw bachelor-file program names to
  canonical ones.
- **Missing input files/fields are meant to be non-fatal in file mode.**
  `files_pipeline._find_first_file()` fuzzy-matches filenames by glob mask
  (e.g. `*асав*.xls*`), and each source loads inside its own try/except with
  a Russian log line on failure, filling the affected columns with 0. Follow
  this pattern for new file sources rather than letting one bad file crash
  the whole run.
- **Log messages and in-sheet labels are in Russian** (internal tool for a
  Russian-speaking team) — match that in new print statements or dashboard
  columns; code identifiers (function/variable names) stay in English.
- **Absolute cell ranges/gaps in `update_pipeline.py`** (e.g.
  `dashboard.update_acell('B50', ...)`, hardcoded `history_data.loc[2026, ...]`
  lookups) are known fragility — several TODOs call this out. Don't add more
  of these without need.
- **`TODO.md`** tracks planned features and technical debt in Russian/English
  shorthand — check it before starting work to avoid duplicating planned
  changes, and update it when you finish something listed there.

## Known gaps (as of this writing — verify before relying on these)

- Several files under `tests/` import things that no longer exist after the
  Bitrix-API refactor: `test_process_helpers.py` imports from a `process`
  module that was renamed to `general_pipeline.py`/`files_pipeline.py`;
  `test_contracts.py` imports from a `contracts` module that is now
  `bitrix_contracts.py`; `test_bitrix.py` and `test_bitrix_pipeline.py`
  reference function names (`collect_portal_360_deals_dataframe`,
  `get_deal_category_id`, `normalize_bitrix_admissions_data`,
  `apply_bitrix_metrics_to_dashboard`) that don't match the current
  (underscore-prefixed, differently-named) functions in `bitrix.py` /
  `bitrix_pipeline.py`. `test_bitrix_pipeline.py` even has a `# TODO repair
  test` comment acknowledging this. Only `test_contracts_summary.py` is
  currently aligned with its module. If you touch a pipeline module, check
  whether its test needs the same repair rather than assuming it's already
  covered.
- `README.md`'s Bitrix usage example (`from bitrix import
  collect_deals_dataframe`) references a function that isn't in `bitrix.py`
  anymore — the real entry points are `create_bitrix_client()` +
  `collect_bitrix_item_sources()`. Don't trust the README's code sample as-is.
- Several dashboard calculations are commented out or stubbed with fixed
  numbers pending a real fix — e.g. `_finalize_dashboard_calculations`'s
  income totals in `bitrix_pipeline.py`, and
  `history_data.loc[2026, 'early_invitations_unique'] = 1658` (hardcoded, `#
  TODO исправить на расчет`). Check the surrounding TODO comment before
  trusting a number that looks suspiciously fixed.

## Writing code in this repo

- **Reuse before adding.** Look for an existing constant, helper, or pattern
  (`col_names.py`, `general_pipeline.insert_values`/`process_by_week`, the
  per-source try/except shape in `files_pipeline.py`, the `BitrixEntity`
  contract shape in `bitrix_contracts.py`) before writing a new one. Don't
  introduce a new column-name string, a new helper duplicating an existing
  one, or a new abstraction layer unless the existing ones genuinely can't do
  the job.
- **Keep changes short and clear.** Prefer the smallest diff that fixes the
  problem over a rewrite; this is a small script-style codebase, not a
  framework — avoid adding new classes, config layers, or indirection for
  their own sake.
