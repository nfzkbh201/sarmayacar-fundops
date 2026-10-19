# Sarmayacar Quarterly Report Importer

Private Figma plugin for turning the FundOps Streamlit `Quarterly Report` data
pack into Sarmayacar-style quarterly report frames.

The source PDF is a visual reference only. The plugin recreates editable Figma
frames and fills them from reviewed workbook data.

## Setup

1. Open Figma Desktop.
2. Go to `Plugins > Development > Import plugin from manifest...`.
3. Select `figma_quarterly_report_plugin/manifest.json`.
4. Run `Sarmayacar Quarterly Report Importer`.
5. Pair the plugin with the private service using the URL and pairing code from
   Streamlit's `Figma handoff` panel.
6. Leave `Receive reports from Streamlit` checked. In Streamlit, click `Send to Figma`.
7. The plugin updates the report and sends completion status back to Streamlit.

See [connection setup](../figma_bridge/README.md) for local testing and hosting.
For file import, expand `File import`, choose the JSON or paste it, and click
`Import report`. Creating template frames first is no longer necessary.

## What The Plugin Creates

- Cover
- Contents
- Highlights
- Three portfolio summary pages in the approved order on every import
- One filled company page per included company
- Filled highlights text when a `Quarterly Commentary` sheet is present
- Filled operator angel pages from the commentary sheet
- Filled financial exhibit pages from the uploaded financial exhibits workbook

## Company Page Data Layers

The generated company template contains these editable layer names:

- `company.name`
- `company.description`
- `company.update_1`
- `company.update_2`
- `company.update_3`
- `company.update_4`
- `notes.disclaimer`

Metric table placeholders:

- `metric_month_1.label` through `metric_month_6.label`
- `metric_1.label`
- `metric_1.month_1` through `metric_1.month_6`
- repeat through `metric_10.*`

New packs supply company-specific financial and operating metric sections with
Actual/Forecast columns. Editable layers use `metrics.section_N.row_N.*` names;
the ten legacy placeholder rows are removed when these sections are present.

Imports now tag their generated frames. Reimporting the same quarter replaces
those frames, including manual edits inside them; older untagged imports and
other quarters remain untouched. The plugin builds the replacement first and
keeps the existing report if rendering fails. No extra Figma file or Page is
created by the connection.

## Fixed Portfolio Layout

The three summary pages always follow Highlights, before company detail pages:

- Printed page 03: Trukker, SimPaisa, Abhi, Bykea, Oladoc, TapMad, Revolving Games,
  OneLoad, Jiye Technologies, Roomy; carrying-value footnote below the table.
- Printed page 04: Procheck, Patari, Dot & Line, Dawaai, Jugnu, Buyzilla; portfolio
  subtotal; Operator Angel Companies heading; Taraki, Savvy Technologies, Delsys Technologies.
- Printed page 05: Startup Early (Holistiq), Scholar Den, House Call, Orko,
  L.L.M.Bots, Aabshar, Raptr Games; grand total.

All three pages share nine column widths, green headers, alternating row shading,
bold company names and totals, centered financial cells, and footer placement.
Rows expand to fit wrapped text. Content that exceeds the page boundary stops the
import and leaves the previous report intact. Missing companies retain their row
with blank data (`-`); unknown or duplicate companies require an explicit layout
review instead of being silently omitted or moving the approved page breaks.

This layout does not calculate or approve numbers. Supplied descriptions, dates,
ownership, board seats, status labels, financial strings and totals pass through
unchanged. Numerical mapping and reconciliation remain a separate review.

## Portfolio Number Sources

The Streamlit data pack maps the selected quarter's financial exhibits, keeping
all 26 summary rows independently of the exhibit display row limit:

- Cost: `SV Total Investment` in the investment schedule, in USD.
- Ownership: `SV Total Ownership` in that schedule, displayed to one decimal percent.
- Carrying value: the matching quarter's Balance Sheet company row, converted from
  USD thousands to whole USD before display rounding.
- Unrealized gain/loss: current carrying value less cost. Old valuation and
  unrealized-gain schedule columns are not used.
- Realized gain/loss: an explicit company `Realized Gain/Loss ($)` column when supplied.
  Without it, display `-`, not an assumed zero or a fund-level income statement figure.

Descriptions, first investment dates and board seats retain the approved report
details. Investment-date conflicts are flagged. Workbook ownership replaces the
approved percentage; blank, invalid or duplicate inputs show `-` with a warning.
Each mapped value carries its source cell in the data pack. Missing amounts
withhold the affected subtotal/total instead of producing a partial sum.

The portfolio review checks grand cost and carrying value against the selected
quarter's Financial Summary control totals. A different quarter's unversioned
cost/ownership schedule cannot be paired with a historical balance column.

## Company Table Rules

`COMPANY_METRIC_SECTION_MAP` in `quarterly_report.py` fixes the row order, source
block, currency and scaling separately for each of the eleven active companies.
SimPaisa uses the USD block and includes interest in Total Income; Tapmad and
OneLoad use their already-scaled reporting blocks. Procheck and Dot & Line remain
in PKR thousands. Counts, occupancy percentages and per-trip values have separate
formats. Each Actual/Forecast value records its source worksheet and cell.
Missing rows remain visible as dashes and raise review warnings, never zeroes.

The renderer uses Helvetica when available, with an explicit Inter fallback
warning. Actual/Forecast columns remain visible even when forecasts are missing.
Descriptions, bullets and row labels are measured before laying out the table;
overflow rejects the update while preserving the existing report. Company order,
contents entries and printed page numbers follow the generated report.

Company-table checks: `python -m unittest discover -s tests -p 'test_company_tables.py'`.
Plugin/transfer checks: `node --test tests/figma_import.test.cjs` (requires the
project Python environment at `.venv`). For local visual checks with Playwright
and Chrome installed, run `node tests/preview_company_tables.cjs DATA_PACK OUTPUT_DIR`.
The preview executes the actual renderer with a measured browser test double,
saves every company page and reports text overflow; it is not a substitute for
the final import check inside Figma.
