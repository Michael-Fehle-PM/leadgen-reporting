# Leadgen Reporting Suite

Python scripts and Oracle SQL for automated lead generation reporting across two business units – military.com (MIL) and Fastweb (FW).

The suite covers two reporting cadences:

- **Daily operations** – how many leads were collected and sent each day, and where's the gap?
- **Partner performance** – how is each partner tracking MTD, versus last month, and versus the last 30 days?

Each script connects to an Oracle database, runs a parameterised SQL query, and produces a formatted Excel workbook ready for distribution.

---

## Files

| File | Type | Description |
|---|---|---|
| `mil_daily_report.py` | Python | Daily leads report for military.com – one tab per date |
| `fw_daily_report.py` | Python | Daily leads report for Fastweb – one tab per date |
| `partner_leads_report.py` | Python | Partner performance report – MIL tab, FW tab, and combined Rollup |
| `query_mil_daily.sql` | SQL | Oracle query for MIL daily report – parameterised by date range |
| `query_fw_daily.sql` | SQL | Oracle query for FW daily report – parameterised by date range |
| `query_mil.sql` | SQL | Oracle query for MIL partner performance – MTD, prior month, last 30 days |
| `query_fw.sql` | SQL | Oracle query for FW partner performance – MTD, prior month, last 30 days |

---

## Script 1 & 2 – Daily leads reports

`mil_daily_report.py` and `fw_daily_report.py` produce a multi-tab Excel workbook with one tab per date in the requested range.

**Each tab contains:**

*Left – Daily summary (cols A–D)*

| Column | Description |
|---|---|
| Date | Report date |
| Collected | Leads received into the system (UTC → ET converted) |
| Sent | Total leads distributed to partners |
| Δ (Delta) | Sent minus Collected – negative values flagged in red |

*Right – Partner breakdown (cols F–L)*

| Column | Description |
|---|---|
| Partner ID / Name | Partner identifier and display name |
| Leads Sent | Leads sent to this partner on this date |
| Cost Per Lead | Current contracted CPL |
| Scrub Rate | Contracted bad lead percentage |
| Max Revenue | Leads Sent × CPL |
| Min Revenue | Leads Sent × (1 − Scrub Rate) × CPL |

Max and Min Revenue are written as live Excel formulas, so the figures update if CPL or scrub rate is adjusted manually after export.

**Usage:**

```bash
# Default: May 1 of current year through yesterday
python mil_daily_report.py
python fw_daily_report.py

# Custom date range
python mil_daily_report.py --start 2026-05-01 --end 2026-05-14
python fw_daily_report.py  --start 2026-05-01 --end 2026-05-14

# Override UTC offset (default: auto-detected)
python mil_daily_report.py --utc-offset -5
```

---

## Script 3 – Partner performance report

`partner_leads_report.py` produces a single workbook with three tabs:

**Tab 1 – MIL Partners**

Full breakdown including leads by traffic source (Google, Bing, Organic) and by product version (V1, V2), across three time windows: MTD, prior month, and last 30 days. Includes a forecast column that extrapolates current MTD run rate to a full-month projection.

**Tab 2 – FW Partners**

Simplified view: MTD, prior month, last 30 days, forecast, CPL, and scrub rate. Fastweb leads do not carry Google/Bing click IDs so source breakdown is not applicable.

**Tab 3 – Rollup**

Combined view of all MIL and FW partners sorted alphabetically, with BU column (MIL/FW) and colour-coded rows. Useful for a single-page view across the full partner portfolio.

**Usage:**

```bash
# Live Oracle connection (default)
python partner_leads_report.py

# CSV mode – useful when no direct database access is available
# Export query results to mil_results.csv and fw_results.csv first
python partner_leads_report.py --csv
```

---

## The SQL

### Daily queries (`query_mil_daily.sql`, `query_fw_daily.sql`)

Both daily queries share a common structure using Oracle CTEs:

- **`date_spine`** – generates one row per date in the requested range using `CONNECT BY`, ensuring dates with zero activity appear in output rather than being silently omitted
- **`partner_contract`** – retrieves the most recent active contract per partner using `ROW_NUMBER() OVER (PARTITION BY PARTNER_ID ORDER BY START_DATE DESC)`, correctly handling partners whose contract rates have changed over time
- **`collected_daily`** – counts leads received per day, converting `CREATED_DATE` from UTC to Eastern Time using the injected `utc_offset` parameter
- **`sent_daily`** – counts leads sent per day per partner, filtered by `SEND_FLAG=1`
- **`sent_daily_total`** – aggregates daily sent totals across all partners for the summary columns

The final `SELECT` cross-joins the date spine with the active partner list and left-joins the activity CTEs, so every (date, partner) combination appears even on days with zero sends.

Parameters are injected as Oracle bind variables at runtime:

```sql
:start_date   -- e.g. DATE '2026-05-01'
:end_date     -- e.g. DATE '2026-05-14'
:utc_offset   -- e.g. -4 (EDT) or -5 (EST)
```

### Performance queries (`query_mil.sql`, `query_fw.sql`)

These run without parameters – all date windows are calculated dynamically from `SYSDATE`.

- **`date_params`** – calculates all date boundaries in one place: MTD start, prior month start/end, last 30 days start/end, days elapsed, and total days in month
- **`agg_mtd` / `agg_prior` / `agg_last30`** – three separate aggregation CTEs, each filtered to its own date window, joined back to the partner list in the final SELECT

The MIL query additionally classifies each lead by traffic source and product version:

```sql
-- Source classification
CASE WHEN ld.GCL_ID     IS NOT NULL THEN 'GOOGLE'
     WHEN ld.MSCLICK_ID IS NOT NULL THEN 'BING'
     ELSE                                'ORGANIC'
END AS lead_source

-- Version classification
CASE WHEN ld.FSRC = 'default-flow'    THEN 'V1'
     WHEN ld.FSRC = 'default-flow-v2' THEN 'V2'
END AS lead_version
```

This attribution logic – Google click ID present → paid search, Microsoft click ID present → Bing, neither → organic – is a standard pattern in martech data pipelines for distinguishing paid from organic lead sources.

---

## UTC to Eastern Time handling

Lead collection timestamps are stored in UTC. The daily scripts automatically determine whether Eastern Time is EDT (UTC−4) or EST (UTC−5) based on US Daylight Saving Time rules:

- DST start: second Sunday in March
- DST end: first Sunday in November

The offset is passed as a bind parameter to the SQL query so the conversion happens in the database. It can be overridden at the command line if needed.

---

## Forecast methodology

The partner performance report includes a forecast column calculated as:

```
Forecast = ROUND( Leads MTD / MAX(Days Elapsed, 1) × Total Days in Month, 0 )
```

This is written as a live Excel formula (`IFERROR(ROUND(...))`) so it recalculates if the underlying data is refreshed. The `MAX(..., 1)` guard prevents division by zero on the first day of the month.

---

## Dependencies

```bash
pip install oracledb openpyxl
```

- `oracledb` – Oracle database connectivity (Python-thin mode, no Oracle Client installation required)
- `openpyxl` – Excel workbook generation

---

## Configuration

Database credentials are stored as constants at the top of each script and have been withheld from this repository. Replace the placeholder values before running:

```python
DB_USER     = "your_username"
DB_PASSWORD = "your_password"
DB_DSN      = "your_host:port/service_name"
```

For production use, credentials should be managed via environment variables or a secrets manager rather than hardcoded constants.

---

## Built with

- Python 3.x
- Oracle Database (via `oracledb` thin client)
- Built with assistance from [Claude](https://claude.ai)

---

## About

Built by [Michael F](https://github.com/MichaelF-PM) during a product management role overseeing lead generation operations at military.com. Background in SaaS product management across martech and fintech, with a focus on data quality, ETL pipelines, and operational tooling.
