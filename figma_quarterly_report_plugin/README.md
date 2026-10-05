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
5. Click `Create XD-style template frames`.
6. In Streamlit, download the Figma data pack from the `Quarterly Report` tab.
7. Paste the JSON data pack into the plugin and click `Import company pages`.

## What The Plugin Creates

- Cover
- Contents
- Highlights
- Portfolio table placeholder
- Company page template
- Operator angel company template
- Financial exhibits placeholder
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

The current Streamlit data pack fills available metrics and leaves extra rows
blank. Portfolio tables are still placeholders for a later phase; highlights,
operator angel pages, company pages, and financial exhibit pages are exported
as editable Figma frames.
