# Daily STN Operations Dashboard

Build a standalone HTML dashboard from the latest STN and optional STR extracts in `This_folder_has_daily_data/`.

## Daily refresh

1. Add that day's STN and STR extract to `This_folder_has_daily_data/`. Files can be CSV, `.xlsx`, or `.xlsm`; use names beginning with `STN ID` and `STR ID`, followed by a date such as `27 SEP 2026`.
2. Install the Excel reader once: `python -m pip install -r requirements.txt`.
3. Run `python build_dashboard.py`.
4. Attach `daily_dashboard.html` to the email. It has no external assets or network dependencies.

The builder selects the newest date in the filename for each extract type. It streams large CSV files and reads the first worksheet of Excel workbooks. To choose other folders or an output path, pass `--data-dir` and `--output`.

## Data interpretation

- STN totals and breakdowns use distinct STN IDs; STN item count uses distinct STN Item IDs (or row count if the IDs are absent).
- Requested and scheduled quantities are summed from the STN item rows. Fulfilment is scheduled quantity divided by requested quantity.
- The supplied STR file's `STR ID` values do not match the STN file's `STN ID` values. The creator chart therefore reports distinct STR IDs per user, not STNs per user. The status-by-user matrix is unavailable until the extracts provide a shared STN identifier; the report calls this out instead of implying a match.
- Figures are calculated from the current daily extracts only. The builder does not retain history or calculate day-over-day trends.# IWIT_daily_dashboard_creation