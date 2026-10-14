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

New packs can supply multiple metric sections with Actual/Forecast columns.
All report sections are editable Figma frames. Portfolio fields and company
metrics still require the separate mapping validation work.

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
