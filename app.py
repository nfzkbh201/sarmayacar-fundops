import os
import sys
from datetime import datetime
from io import BytesIO
from pathlib import Path

import streamlit as st
import pandas as pd

APP_ROOT = Path(__file__).resolve().parent

monthly_reporting_root = Path(
    os.environ.get(
        "MONTHLY_REPORTING_ROOT",
        APP_ROOT,
    )
).resolve()
if str(monthly_reporting_root) not in sys.path:
    sys.path.insert(0, str(monthly_reporting_root))

from monthly_reporting_page import render_monthly_reporting_page

st.set_page_config(
    page_title="Sarmayacar FundOps Dashboard",
    layout="wide"
)

st.title("Sarmayacar FundOps Dashboard")
st.write("Upload updated files from the sidebar, or use the default project files.")

investor_master = APP_ROOT / "Investor_Master" / "Copy of Investor_Details.xlsx"
drawdown_file = APP_ROOT / "Inputs" / "SV-Drawdown (Mar 2026).XLSX"
contact_file = Path(
    os.environ.get(
        "FUNDOPS_CONTACT_FILE",
        APP_ROOT / "Investor_Master" / "SV contact list.xlsx",
    )
).resolve()
template_file = APP_ROOT / "Templates" / "Call_Notice_Template.docx"
distribution_file = APP_ROOT / "Inputs" / "Sarmayacar Distribution Payments (2026)-2.xlsx"
generated_notices_root = Path(
    os.environ.get("FUNDOPS_OUTPUT_ROOT", APP_ROOT / "Generated_Notices")
).resolve()

st.sidebar.header("Data Sources")
uploaded_investor_master = st.sidebar.file_uploader(
    "Investor master",
    type=["xlsx"],
    key="investor_master_upload",
)
uploaded_drawdown_file = st.sidebar.file_uploader(
    "Drawdown file",
    type=["xlsx"],
    key="drawdown_file_upload",
)
uploaded_contact_file = st.sidebar.file_uploader(
    "Contact list",
    type=["xlsx"],
    key="contact_file_upload",
)
uploaded_template_file = st.sidebar.file_uploader(
    "Call notice template",
    type=["docx"],
    key="template_file_upload",
)
drawdown_sheet_name = st.sidebar.text_input(
    "Drawdown sheet name",
    value="Call Notices Mar 2026",
)
drawdown_header_row = st.sidebar.number_input(
    "Drawdown header row",
    min_value=1,
    value=7,
)


def read_excel_source(uploaded_file, default_path, **kwargs):
    if uploaded_file is not None:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, **kwargs)
    return pd.read_excel(default_path, **kwargs)


def excel_sheet_names(uploaded_file, default_path):
    if uploaded_file is not None:
        uploaded_file.seek(0)
        return pd.ExcelFile(uploaded_file).sheet_names
    return pd.ExcelFile(default_path).sheet_names


def read_contact_source(uploaded_file, default_path):
    old_contact_columns = [
        "Member Name",
        "Call Notice Contact Name",
        "Call Notice Contact Email",
        "Postal Address",
    ]
    master_contact_columns = {
        "Limited Partner": "Member Name",
        "Principal Contact Name": "Call Notice Contact Name",
        "Email Address": "Call Notice Contact Email",
        "Address": "Postal Address",
    }

    if uploaded_file is not None:
        uploaded_file.seek(0)
        source = uploaded_file
    else:
        source = default_path

    excel_file = pd.ExcelFile(source)
    for sheet_name in excel_file.sheet_names:
        for header_row in range(0, 8):
            if uploaded_file is not None:
                uploaded_file.seek(0)
            candidate_df = pd.read_excel(
                source,
                sheet_name=sheet_name,
                header=header_row,
            )
            candidate_df.columns = [str(c).strip() for c in candidate_df.columns]

            if all(col in candidate_df.columns for col in old_contact_columns):
                return candidate_df

            if all(col in candidate_df.columns for col in master_contact_columns):
                return candidate_df.rename(columns=master_contact_columns)

    if uploaded_file is not None:
        uploaded_file.seek(0)
    fallback_df = pd.read_excel(source)
    fallback_df.columns = [str(c).strip() for c in fallback_df.columns]
    return fallback_df


def make_document(uploaded_file, default_path):
    if uploaded_file is not None:
        return Document(BytesIO(uploaded_file.getvalue()))
    return Document(default_path)


def source_available(uploaded_file, default_path):
    return uploaded_file is not None or Path(default_path).exists()


def source_label(uploaded_file, default_path):
    if uploaded_file is not None:
        return f"Uploaded: {uploaded_file.name}"
    if Path(default_path).exists():
        return f"Default: {default_path}"
    return f"Upload needed: {Path(default_path).name}"


def missing_source_names(required_sources):
    return [
        label
        for label, uploaded_file, default_path in required_sources
        if not source_available(uploaded_file, default_path)
    ]


def is_missing(value):
    if value is None:
        return True
    text = str(value).strip()
    return text == "" or text.lower() == "nan"


def clean_filename_part(value):
    text = str(value).strip()
    for char in ['<', '>', ':', '"', '/', '\\', '|', '?', '*']:
        text = text.replace(char, "-")
    return " ".join(text.split())


manual_investor_display_name_map = {
    "Hermann Futter": "Compass Gruppe GmbH",
}


st.sidebar.caption(source_label(uploaded_investor_master, investor_master))
st.sidebar.caption(source_label(uploaded_drawdown_file, drawdown_file))
st.sidebar.caption(source_label(uploaded_contact_file, contact_file))
st.sidebar.caption(source_label(uploaded_template_file, template_file))

capital_calls_tab, distribution_notices_tab, monthly_reporting_tab = st.tabs(
    ["Capital Calls", "Distribution Notices", "Monthly Reporting"]
)

with capital_calls_tab:
    st.success("Data sources are ready.")

    required_capital_sources = [
        ("Investor master", uploaded_investor_master, investor_master),
        ("Drawdown file", uploaded_drawdown_file, drawdown_file),
        ("Contact list", uploaded_contact_file, contact_file),
        ("Call notice template", uploaded_template_file, template_file),
    ]
    missing_capital_sources = missing_source_names(required_capital_sources)

    if missing_capital_sources:
        st.info(
            "Upload these files from the sidebar to use Capital Calls: "
            + ", ".join(missing_capital_sources)
            + "."
        )
    else:
        st.success("Required files are available.")

        investor_df = read_excel_source(uploaded_investor_master, investor_master, header=5)
        investor_df = investor_df.astype(str)
        investor_df.columns = [str(c).strip() for c in investor_df.columns]

        required_investor_columns = [
            "Limited Partner",
            "Investor_ID",
            "Payment_Reference",
        ]
        missing_investor_columns = [
            col for col in required_investor_columns if col not in investor_df.columns
        ]
        if missing_investor_columns:
            st.error(
                "Investor master is missing: "
                + ", ".join(missing_investor_columns)
            )
            st.stop()

        st.subheader("Investor Master Preview")
        st.write(investor_df.head())

        st.write("Columns found:")

        investor_col = "Limited Partner"
        excluded_investor_names = {"Sub Total", "Total", "Grand Total"}

        investor_names = (
            investor_df[investor_col]
            .dropna()
            .astype(str)
            .str.strip()
            .loc[lambda names: ~names.isin(excluded_investor_names)]
            .sort_values()
            .unique()
            .tolist()
        )

        st.subheader("Select Investors")

        if "selected_investors" not in st.session_state:
            st.session_state["selected_investors"] = []

        st.session_state["selected_investors"] = [
            investor
            for investor in st.session_state["selected_investors"]
            if investor in investor_names
        ]

        selection_col_1, selection_col_2 = st.columns(2)
        with selection_col_1:
            if st.button("Select All Investors"):
                st.session_state["selected_investors"] = investor_names
        with selection_col_2:
            if st.button("Clear Selection"):
                st.session_state["selected_investors"] = []

        selected_investors = st.multiselect(
            "Choose investors for this capital call:",
            investor_names,
            key="selected_investors",
            format_func=lambda investor: manual_investor_display_name_map.get(
                investor,
                investor,
            ),
        )

        st.write(f"Selected investors: {len(selected_investors)}")



        st.subheader("Drawdown Preview")

        try:
            drawdown_df = read_excel_source(
                uploaded_drawdown_file,
                drawdown_file,
                sheet_name=drawdown_sheet_name,
                header=int(drawdown_header_row) - 1,
            )
        except ValueError:
            st.error(f'Drawdown sheet not found: "{drawdown_sheet_name}"')
            st.stop()

        required_drawdown_columns = [
            "Limited Partner",
            "Total Commitment",
            "Pro-Rata Distribution",
            "Current Drawdown",
            "Total Amount",
            "Total % Drawn",
        ]

        drawdown_df.columns = [str(c).strip() for c in drawdown_df.columns]
        missing_drawdown_columns = [
            col for col in required_drawdown_columns if col not in drawdown_df.columns
        ]
        if missing_drawdown_columns:
            st.error(
                "Drawdown file is missing: "
                + ", ".join(missing_drawdown_columns)
            )
            st.stop()

        contact_df = read_contact_source(uploaded_contact_file, contact_file)

        required_contact_columns = [
            "Member Name",
            "Call Notice Contact Name",
            "Call Notice Contact Email",
            "Postal Address",
        ]
        missing_contact_columns = [
            col for col in required_contact_columns if col not in contact_df.columns
        ]
        if missing_contact_columns:
            st.error(
                "Contact list is missing: "
                + ", ".join(missing_contact_columns)
            )
            st.stop()

        if uploaded_template_file is not None and not uploaded_template_file.name.lower().endswith(".docx"):
            st.error("Call notice template must be a .docx file.")
            st.stop()

        st.caption(
            "Using "
            + source_label(uploaded_investor_master, investor_master)
            + " | "
            + source_label(uploaded_drawdown_file, drawdown_file)
            + " | "
            + source_label(uploaded_contact_file, contact_file)
            + " | "
            + source_label(uploaded_template_file, template_file)
        )

        drawdown_df = drawdown_df.dropna(subset=["Limited Partner"])

        st.write(drawdown_df[[
            "Limited Partner",
            "Total Commitment",
            "Pro-Rata Distribution",
            "Current Drawdown",
            "Total Amount",
            "Total % Drawn"
        ]].head())

        st.subheader("Notice Dates")

        notice_date = st.date_input("Notice Date")
        value_date = st.date_input("For value on or before")
        notice_period_label = st.text_input(
            "File period label",
            value=notice_date.strftime("%b %Y"),
        )

        st.write("Notice Date selected:", notice_date.strftime("%d %B %Y"))
        st.write("Value Date selected:", value_date.strftime("%d %B %Y"))

        manual_drawdown_name_map = {
            "IFC": "International Finance Corporation (IFC)",
            "IFC (WeFi)": "IFC (WeFi Implementing Partner)",
            "Lucky Group": "GrandCres Investment Ltd",
            "Growbiz Enterprises Ltd (Hilton Pharma)": "Growbiz Enterprises Ltd",
            "SYM Holdings Ltd (Hilton Pharma)": "SYM Holdings Ltd",
            "Yumna Motiwala": "Yumna Jabbar Motiwala",
            "Nasser Ahmad": "Nasser Aziz Ahmad",
            "GP": "Sarmayacar BV",
            "Fuya Holding GmbH": "Yassir Pasha (Previously Fuya Holding GmbH)",
            "Ahmed S. Hameed (RUKS Trust)": "RUKS International Trust",
            "Aleem Siddiqi": "Aleem Hisam Siddiqi",
            "Ali Almakky": "Ali Omar Almakky",
            "Andreas Tuczka (The Aldridge Trust)": "The Aldridge Trust",
            "Cedric Koehler": "Cedric Sebastian Kohler",
            "Faisal": "Faisal Aziz Essa",
            "Fawad Zakariya": "Mohammad Fawad Zakariya",
            "Fredrich Jergitsch (Freshfields)": "Friedrich Jergitsch",
            "Hassan Bhatti": "Hassan Waqar Bhatti",
            "Hermann Futter": "Compass Gruppe GmbH",
            "Jahangir": "Jahangir Abdullah Rasheed",
            "Johannes Hornig": "J Hornig Holding GmbH",
            "Kaveh Tehrani": "Seyed Kaveh Hosseini Tehrani",
            "Kishmish Ventures": "Kishmish Ventures ApS",
            "L4 Invest GmbH (Markus Pernusch)": "L4 Invest GmbH",
            "Magdalena Beirder": "Magdalena Biereder",
            "Michael Schernthaner (Pure Performance GmbH)": "Pure Performance GmbH",
            "Oldcastle Limited": "Oldcastle Limited (Mario Altenburger)",
            "Soofian Zuberi": "Soofian J Zuberi",
        }

        manual_contact_name_map = {
            "Aleem Siddiqi": "Aleem Hisam Siddiqi",
            "Ali Almakky": "Ali Omar Almakky",
            "Andreas Tuczka (The Aldridge Trust)": "The Aldridge Trust",
            "Bancroft Ventures LLC": "Bancroft Ventures LLC (Ahmed Saeed Chaudhary)",
            "Cedric Koehler": "Cedric Sebastian Kohler",
            "Faisal": "Faisal Aziz Essa",
            "Fawad Zakariya": "Mohammad Fawad Zakariya",
            "Fredrich Jergitsch (Freshfields)": "Fredrich Jergitsch",
            "Hassan Bhatti": "Hassan Waqar Bhatti",
            "Hermann Futter": "Compass Gruppe GmbH",
            "Jahangir": "Jahangir Abdullah Rasheed",
            "Johannes Hornig": "J Hornig Beteiligungsgesellschaft m.b.H",
            "Kaveh Tehrani": "Seyed Kaveh Hosseini Tehrani",
            "Kishmish Ventures": "Kishmish Ventures ApS",
            "L4 Invest GmbH (Markus Pernusch)": "L4 Invest GmbH",
            "Lucky Group": "GrandCres Investment Ltd",
            "Magdalena Beirder": "Magdalena Biereder",
            "Michael Schernthaner (Pure Performance GmbH)": "Pure Performance GmbH",
            "Mohammad Kashif Rehman": "Muhammad Kashif Rehman",
            "Soofian Zuberi": "Soofian J Zuberi",
            "The Rizvi Family Trust": "The Rizvi Family Trust (Hasan Ghulam Rizvi)",
            "Yumna Motiwala": "Yumna Jabbar Motiwala",
        }

        if selected_investors:
            st.subheader("Selected Investor Drawdown Match Preview")

            preview_rows = []

            for investor in selected_investors:
                clean_investor = investor.strip()
                drawdown_name = manual_drawdown_name_map.get(clean_investor, clean_investor)
                notice_display_name = drawdown_name

                match = drawdown_df[
                    drawdown_df["Limited Partner"].astype(str).str.strip() == drawdown_name
                ]

                if len(match) > 0:
                    row = match.iloc[0]
                    preview_rows.append({
                        "Investor Master Name": investor,
                        "Notice Display Name": notice_display_name,
                        "Drawdown Match Name": drawdown_name,
                        "Current Drawdown": row["Current Drawdown"],
                        "Total Amount": row["Total Amount"],
                        "Total % Drawn": row["Total % Drawn"],
                        "Status": "Matched"
                    })
                else:
                    preview_rows.append({
                        "Investor Master Name": investor,
                        "Notice Display Name": notice_display_name,
                        "Drawdown Match Name": drawdown_name,
                        "Current Drawdown": "",
                        "Total Amount": "",
                        "Total % Drawn": "",
                        "Status": "NOT MATCHED"
                    })

            st.write(pd.DataFrame(preview_rows))

        if selected_investors:
            st.subheader("Merged Investor Data Preview")

            merged_rows = []
            generation_issues = []

            for investor in selected_investors:
                clean_investor = investor.strip()
                drawdown_name = manual_drawdown_name_map.get(clean_investor, clean_investor)
                notice_display_name = drawdown_name

                investor_match = investor_df[
                    investor_df["Limited Partner"].astype(str).str.strip() == clean_investor
                ]

                contact_name = manual_contact_name_map.get(clean_investor, clean_investor)

                contact_match = contact_df[
                    contact_df["Member Name"].astype(str).str.strip() == contact_name
                ]

                drawdown_match = drawdown_df[
                    drawdown_df["Limited Partner"].astype(str).str.strip() == drawdown_name
                ]

                merged_row = {
                    "Investor": clean_investor,
                    "Notice Display Name": notice_display_name,
                    "Drawdown Name": drawdown_name,
                    "Investor ID": investor_match.iloc[0]["Investor_ID"] if len(investor_match) else "",
                    "Payment Reference": investor_match.iloc[0]["Payment_Reference"] if len(investor_match) else "",
                    "Contact Name": contact_match.iloc[0]["Call Notice Contact Name"] if len(contact_match) else "",
                    "Email": contact_match.iloc[0]["Call Notice Contact Email"] if len(contact_match) else "",
                    "Postal Address": contact_match.iloc[0]["Postal Address"] if len(contact_match) else "",
                    "Current Drawdown": drawdown_match.iloc[0]["Current Drawdown"] if len(drawdown_match) else "",
                    "Total Amount": drawdown_match.iloc[0]["Total Amount"] if len(drawdown_match) else "",
                    "Total % Drawn": drawdown_match.iloc[0]["Total % Drawn"] if len(drawdown_match) else "",
                }
                merged_rows.append(merged_row)

                missing_fields = [
                    field for field, value in merged_row.items()
                    if field not in {"Investor", "Notice Display Name", "Drawdown Name"} and is_missing(value)
                ]

                if len(investor_match) == 0:
                    missing_fields.append("Investor master match")
                if len(contact_match) == 0:
                    missing_fields.append("Contact list match")
                if len(drawdown_match) == 0:
                    missing_fields.append("Drawdown match")

                if missing_fields:
                    generation_issues.append({
                        "Investor": clean_investor,
                        "Issue": ", ".join(dict.fromkeys(missing_fields)),
                    })

            st.write(pd.DataFrame(merged_rows))

            st.subheader("Pre-Generation Check")
            if generation_issues:
                st.error(f"{len(generation_issues)} investor(s) need fixes before generation.")
                st.write(pd.DataFrame(generation_issues))
            else:
                st.success("All selected investors are ready to generate.")


        if selected_investors:
            can_generate = "generation_issues" in locals() and len(generation_issues) == 0
            if st.button("Generate Selected Notices", disabled=not can_generate):
                from docx import Document
                from docx.shared import Pt, Inches, RGBColor
                from docx.oxml import OxmlElement
                from docx.oxml.ns import qn
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                from docx.enum.table import WD_TABLE_ALIGNMENT
                from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

                generated_notices_root.mkdir(parents=True, exist_ok=True)
                batch_folder_name = datetime.now().strftime("%Y-%m-%d_%H-%M-%S_Capital_Call")
                batch_output_dir = generated_notices_root / batch_folder_name
                batch_output_dir.mkdir(parents=True, exist_ok=True)
                filename_period = clean_filename_part(notice_period_label) or notice_date.strftime("%b %Y")

                def set_run_font(run, font_name="Calibri", size=11, bold=False):
                    run.font.name = font_name
                    run.font.size = Pt(size)
                    run.bold = bold

                def set_paragraph_calibri_11(paragraph):
                    for run in paragraph.runs:
                        set_run_font(run, "Calibri", 11)

                def replace_text(doc, old, new):
                    for p in doc.paragraphs:
                        if old in p.text:
                            p.text = p.text.replace(old, str(new))
                            set_paragraph_calibri_11(p)

                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                if old in cell.text:
                                    cell.text = cell.text.replace(old, str(new))
                                    for p in cell.paragraphs:
                                        set_paragraph_calibri_11(p)

                def split_address(address):
                    parts = [x.strip() for x in str(address).split(",") if x.strip()]
                    while len(parts) < 5:
                        parts.append("")
                    return parts[:5]

                def fix_signature_date_line(doc, notice_date_text):
                    for paragraph in doc.paragraphs:
                        text = paragraph.text.strip()
                        if text.startswith("Date:") and text.endswith("Date:"):
                            p_pr = paragraph._p.get_or_add_pPr()
                            old_tabs = p_pr.find(qn("w:tabs"))
                            if old_tabs is not None:
                                p_pr.remove(old_tabs)

                            tabs = OxmlElement("w:tabs")
                            for position in (720, 4320):
                                tab = OxmlElement("w:tab")
                                tab.set(qn("w:val"), "left")
                                tab.set(qn("w:pos"), str(position))
                                tabs.append(tab)
                            p_pr.append(tabs)

                            paragraph.text = f"Date:\t{notice_date_text}\tDate:"
                            set_paragraph_calibri_11(paragraph)

                def clear_old_schedule(doc):
                    # Remove old template tables from page 3
                    for table in list(doc.tables):
                        tbl = table._element
                        tbl.getparent().remove(tbl)

                    # Remove only drawings/images from Schedule I onwards.
                    # This preserves the signature image above Rabeel Warraich.
                    schedule_started = False
                    for p in doc.paragraphs:
                        if "Schedule I to the Call Notice" in p.text:
                            schedule_started = True

                        if schedule_started:
                            for run in list(p.runs):
                                if run._element.xpath(".//w:drawing") or run._element.xpath(".//w:pict"):
                                    run._element.getparent().remove(run._element)

                def shade_cell(cell, fill):
                    tc_pr = cell._tc.get_or_add_tcPr()
                    shd = OxmlElement("w:shd")
                    shd.set(qn("w:fill"), fill)
                    tc_pr.append(shd)

                def set_table_indent(table, indent_twips=1080):
                    tbl_pr = table._tbl.tblPr
                    tbl_ind = tbl_pr.find(qn("w:tblInd"))
                    if tbl_ind is None:
                        tbl_ind = OxmlElement("w:tblInd")
                        tbl_pr.append(tbl_ind)
                    tbl_ind.set(qn("w:w"), str(indent_twips))
                    tbl_ind.set(qn("w:type"), "dxa")

                def add_drawdown_table(doc, row):
                    from docx.shared import Pt, Inches, RGBColor
                    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
                    from docx.enum.table import WD_TABLE_ALIGNMENT
                    from docx.oxml import OxmlElement
                    from docx.oxml.ns import qn

                    # Remove old duplicate paragraph headings only
                    for para in list(doc.paragraphs):
                        if para.text.strip() == "Drawdown Schedule (Table 1)":
                            para._element.getparent().remove(para._element)

                    def get_value(row, cols):
                        for col in cols:
                            if col in row.index:
                                val = row.get(col)
                                if val is not None and str(val).strip() != "" and str(val).lower() != "nan":
                                    return val
                        return ""

                    def to_float(value):
                        try:
                            s = str(value).replace("$", "").replace(",", "").replace("%", "").strip()
                            if s == "" or s.lower() == "nan":
                                return None
                            return float(s)
                        except Exception:
                            return None

                    def fmt_money(value, decimals=0):
                        num = to_float(value)
                        if num is None:
                            return ""
                        return "$ {:,.{}f}".format(num, decimals)

                    def fmt_percent(value):
                        num = to_float(value)
                        if num is None:
                            return ""
                        if num <= 1:
                            num *= 100
                        if 0 < abs(num) < 1:
                            return f"{num:.2f}%"
                        return f"{num:.0f}%"

                    limited_partner = get_value(row, ["Limited Partner", "Drawdown Name", "Investor"])
                    total_commitment = get_value(row, ["Total Commitment"])
                    pro_rata = get_value(row, ["Pro-Rata Distribution", "Pro Rata Distribution", "Pro-Rata"])

                    amount_prev = get_value(row, [
                        "Amount Previously Drawn of Total Commitment",
                        "Amount Previously Drawn",
                        "Previously Drawn",
                        "Previous Drawdown",
                        "Amount Previously Drawn of Total\nCommitment",
                    ])

                    percent_prev = get_value(row, [
                        "% Previously Drawn of Total Commitment",
                        "% Previously Drawn",
                        "Previously Drawn %",
                        "% Previously Drawn of Total\nCommitment",
                    ])

                    current_drawdown = get_value(row, ["Current Drawdown"])
                    total_amount = get_value(row, ["Total Amount", "Total Amount Drawn", "Total Drawn"])
                    total_percent = get_value(row, ["Total % Drawn", "Total Percent Drawn"])

                    # Fallback calculations
                    if amount_prev == "":
                        total_amount_num = to_float(total_amount)
                        current_drawdown_num = to_float(current_drawdown)
                        if total_amount_num is not None and current_drawdown_num is not None:
                            amount_prev = total_amount_num - current_drawdown_num

                    if total_amount == "":
                        amount_prev_num = to_float(amount_prev)
                        current_drawdown_num = to_float(current_drawdown)
                        if amount_prev_num is not None and current_drawdown_num is not None:
                            total_amount = amount_prev_num + current_drawdown_num

                    if percent_prev == "":
                        amount_prev_num = to_float(amount_prev)
                        total_commitment_num = to_float(total_commitment)
                        if amount_prev_num is not None and total_commitment_num:
                            percent_prev = amount_prev_num / total_commitment_num

                    if total_percent == "":
                        total_amount_num = to_float(total_amount)
                        total_commitment_num = to_float(total_commitment)
                        if total_amount_num is not None and total_commitment_num:
                            total_percent = total_amount_num / total_commitment_num

                    headers = [
                        "Limited\nPartner",
                        "Total\nCommitment",
                        "Pro-Rata\nDistribution",
                        "Amount\nPreviously\nDrawn",
                        "% Previously\nDrawn of Total\nCommitment",
                        "Current\nDrawdown",
                        "Total\nAmount\nDrawn",
                        "Total %\nDrawn",
                    ]

                    values = [
                        str(limited_partner),
                        fmt_money(total_commitment, 0),
                        fmt_percent(pro_rata),
                        fmt_money(amount_prev, 0),
                        fmt_percent(percent_prev),
                        fmt_money(current_drawdown, 2),
                        fmt_money(total_amount, 0),
                        fmt_percent(total_percent),
                    ]

                    width_inches = [1.00, 0.75, 0.75, 0.85, 1.05, 0.80, 0.85, 0.60]
                    widths = [Inches(x) for x in width_inches]

                    table = doc.add_table(rows=3, cols=8)
                    table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.autofit = False

                    for idx, width in enumerate(widths):
                        for cell in table.columns[idx].cells:
                            cell.width = width

                    def shade_cell(cell, fill):
                        tc_pr = cell._tc.get_or_add_tcPr()
                        shd = OxmlElement("w:shd")
                        shd.set(qn("w:fill"), fill)
                        tc_pr.append(shd)

                    def set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
                        tc_pr = cell._tc.get_or_add_tcPr()
                        tc_mar = tc_pr.find(qn("w:tcMar"))
                        if tc_mar is None:
                            tc_mar = OxmlElement("w:tcMar")
                            tc_pr.append(tc_mar)

                        for margin_name, margin_value in {
                            "top": top,
                            "start": start,
                            "bottom": bottom,
                            "end": end,
                        }.items():
                            margin = tc_mar.find(qn(f"w:{margin_name}"))
                            if margin is None:
                                margin = OxmlElement(f"w:{margin_name}")
                                tc_mar.append(margin)
                            margin.set(qn("w:w"), str(margin_value))
                            margin.set(qn("w:type"), "dxa")

                    def set_cell_text(cell, text, bold=False, size=8, white=False, align="center"):
                        cell.text = ""
                        p = cell.paragraphs[0]
                        if align == "right":
                            p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                        elif align == "left":
                            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                        else:
                            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(0)
                        r = p.add_run(str(text))
                        r.bold = bold
                        r.font.name = "Calibri"
                        r.font.size = Pt(size)
                        if white:
                            r.font.color.rgb = RGBColor(255, 255, 255)

                    # Row 0: heading merged inside same table, so it aligns exactly
                    heading_cell = table.cell(0, 0).merge(table.cell(0, 7))
                    set_cell_margins(heading_cell, top=0, start=0, bottom=0, end=0)
                    set_cell_text(
                        heading_cell,
                        "Drawdown Schedule (Table 1)",
                        bold=True,
                        size=10,
                        white=False,
                        align="left"
                    )

                    # Row 1: headers
                    for i, header in enumerate(headers):
                        cell = table.cell(1, i)
                        shade_cell(cell, "2E7D67")
                        set_cell_text(cell, header, bold=True, size=8, white=True, align="center")

                    # Row 2: values
                    money_cols = {1, 3, 5, 6}
                    for i, value in enumerate(values):
                        cell = table.cell(2, i)
                        shade_cell(cell, "F2F2F2")
                        align = "right" if i in money_cols else "center"
                        set_cell_text(cell, value, bold=True, size=8, white=False, align=align)


                for investor in selected_investors:
                    clean_investor = investor.strip()
                    drawdown_name = manual_drawdown_name_map.get(clean_investor, clean_investor)
                    notice_display_name = drawdown_name

                    investor_match = investor_df[
                        investor_df["Limited Partner"].astype(str).str.strip() == clean_investor
                    ]

                    contact_name = manual_contact_name_map.get(clean_investor, clean_investor)

                    contact_match = contact_df[
                        contact_df["Member Name"].astype(str).str.strip() == contact_name
                    ]

                    drawdown_match = drawdown_df[
                        drawdown_df["Limited Partner"].astype(str).str.strip() == drawdown_name
                    ]

                    if len(drawdown_match) == 0:
                        st.error(f"No drawdown match found for {investor}")
                        continue

                    inv_row = investor_match.iloc[0] if len(investor_match) else {}
                    contact_row = contact_match.iloc[0] if len(contact_match) else {}
                    draw_row = drawdown_match.iloc[0]

                    address = contact_row.get("Postal Address", "")
                    payment_ref = inv_row.get("Payment_Reference", "A0007713")

                    current_drawdown = float(draw_row["Current Drawdown"])
                    amount_text = f'US $ {current_drawdown:,.0f}'

                    doc = make_document(uploaded_template_file, template_file)

                    clear_old_schedule(doc)

                    addr1, addr2, addr3, addr4, addr5 = split_address(address)

                    replace_text(doc, "28Lex Capital LLC", notice_display_name)
                    replace_text(doc, "Corporation Trust Center 1209", addr1)
                    replace_text(doc, "Orange Street, Wilmington", addr2)
                    replace_text(doc, "Delaware 19801", addr3)
                    replace_text(doc, "New Castle", addr4)
                    replace_text(doc, "United States of America", addr5)

                    replace_text(doc, "US $ 11,387", amount_text)
                    replace_text(doc, "20th May 2026", value_date.strftime("%d %B %Y"))
                    replace_text(doc, "27th Apr 2026", notice_date.strftime("%d %B %Y"))
                    replace_text(doc, "A0007713", payment_ref)
                    fix_signature_date_line(doc, notice_date.strftime("%d %B %Y"))

                    # Ensure key edited fields are Calibri 11
                    for p in doc.paragraphs:
                        txt = p.text
                        if (
                            notice_display_name in txt
                            or addr1 in txt
                            or addr2 in txt
                            or addr3 in txt
                            or addr4 in txt
                            or addr5 in txt
                            or "Amount of Capital Contribution" in txt
                            or "For value on or before" in txt
                            or "Payment Ref:" in txt
                            or "Date:" in txt
                        ):
                            set_paragraph_calibri_11(p)

                    notice_draw_row = draw_row.copy()
                    notice_draw_row["Limited Partner"] = notice_display_name
                    add_drawdown_table(doc, notice_draw_row)

                    safe_name = clean_filename_part(notice_display_name)
                    output_filename = f"{safe_name} - Call Notice ({filename_period}).docx"
                    output_path = batch_output_dir / output_filename
                    doc.save(output_path)

                st.success(f"Generated {len(selected_investors)} Word notices in {batch_output_dir}.")


with distribution_notices_tab:
    st.subheader("Distribution Notices")
    st.write("Upload the annual distribution file and notice template, then choose the sheet/header row to preview.")

    distribution_upload_col_1, distribution_upload_col_2 = st.columns(2)
    with distribution_upload_col_1:
        uploaded_distribution_file = st.file_uploader(
            "Annual distribution file",
            type=["xlsx"],
            key="distribution_file_upload",
        )
    with distribution_upload_col_2:
        uploaded_distribution_template = st.file_uploader(
            "Distribution notice template",
            type=["docx"],
            key="distribution_template_upload",
        )

    uploaded_distribution_sample = st.file_uploader(
        "Sample generated distribution notice",
        type=["docx"],
        key="distribution_sample_upload",
    )

    st.caption(source_label(uploaded_distribution_file, distribution_file))

    try:
        distribution_sheet_options = excel_sheet_names(uploaded_distribution_file, distribution_file)
    except FileNotFoundError:
        distribution_sheet_options = []
        st.warning("No default distribution file found. Upload an annual distribution file to continue.")

    if distribution_sheet_options:
        default_sheet_index = (
            distribution_sheet_options.index("Jugnu ")
            if "Jugnu " in distribution_sheet_options
            else 0
        )
        distribution_sheet_name = st.selectbox(
            "Distribution sheet",
            distribution_sheet_options,
            index=default_sheet_index,
        )
        default_distribution_header_row = (
            8 if distribution_sheet_name == "Distribution Payments" else 4
        )
        distribution_header_row = st.number_input(
            "Distribution header row",
            min_value=1,
            value=default_distribution_header_row,
            key="distribution_header_row",
        )

        try:
            distribution_df = read_excel_source(
                uploaded_distribution_file,
                distribution_file,
                sheet_name=distribution_sheet_name,
                header=int(distribution_header_row) - 1,
            )
            distribution_df.columns = [str(c).strip() for c in distribution_df.columns]
            distribution_df = distribution_df.dropna(how="all")
        except Exception as exc:
            st.error(f"Could not read distribution sheet: {exc}")
            distribution_df = pd.DataFrame()

        if not distribution_df.empty:
            st.subheader("Distribution Source Preview")
            st.write(distribution_df.head(20))

            st.subheader("Distribution Columns")
            st.write(pd.DataFrame({"Column": distribution_df.columns.tolist()}))

            if "Limited Partner" not in distribution_df.columns:
                st.warning("Limited Partner column was not found. Adjust the sheet/header row before generation.")
            else:
                distribution_workflow = st.radio(
                    "Distribution workflow",
                    ["Jugnu liquidation", "Dividend distribution"],
                    horizontal=True,
                )

                distribution_preview_df = pd.DataFrame()

                if distribution_workflow == "Jugnu liquidation":
                    amount_column_options = [
                        col for col in distribution_df.columns
                        if "Jugnu" in col or "Liquidation" in col
                    ]
                    amount_column_options = amount_column_options or distribution_df.columns.tolist()
                    jugnu_amount_column = st.selectbox(
                        "Distribution amount column",
                        amount_column_options,
                    )

                    phase_columns = [
                        col for col in distribution_df.columns
                        if col.startswith("Unnamed")
                        and distribution_df[col].astype(str).str.contains("Release in phase", case=False, na=False).any()
                    ]
                    phase_column = st.selectbox(
                        "Phase/status column",
                        phase_columns or distribution_df.columns.tolist(),
                    )
                    phase_options = (
                        distribution_df[phase_column]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .loc[lambda values: values != ""]
                        .unique()
                        .tolist()
                    )
                    selected_phases = st.multiselect(
                        "Release phases to include",
                        phase_options,
                        default=phase_options,
                    )
                    total_distribution_amount = st.number_input(
                        "Total proceeds amount",
                        min_value=0.0,
                        value=300000.0,
                        step=1000.0,
                    )
                    distribution_description = st.text_input(
                        "Distribution description",
                        value="Jugnu's (portfolio company) liquidation",
                    )

                    distribution_preview_df = distribution_df[
                        distribution_df[phase_column].astype(str).str.strip().isin(selected_phases)
                    ][[
                        "Limited Partner",
                        "Total Commitment",
                        "Pro-Rata Distribution",
                        jugnu_amount_column,
                        phase_column,
                    ]].copy()
                    distribution_preview_df = distribution_preview_df.rename(columns={
                        jugnu_amount_column: "Distribution Amount",
                        phase_column: "Release Phase",
                    })

                    st.caption(
                        f"Description: {distribution_description}; total proceeds: USD {total_distribution_amount:,.0f}"
                    )

                else:
                    payout_options = {
                        "Payout 1": ("Distribution 1", "Release Date"),
                        "Payout 2": ("Distribution 2", "Release Date.1"),
                        "Payout 3": ("Distribution 3", "Release Date.2"),
                    }
                    selected_payout = st.selectbox(
                        "Payout to generate",
                        list(payout_options.keys()),
                        index=2,
                    )
                    amount_column, release_date_column = payout_options[selected_payout]
                    distribution_preview_df = distribution_df[
                        distribution_df[amount_column].notna()
                    ][[
                        "Limited Partner",
                        "Total Commitment",
                        "Pro-Rata Distribution",
                        amount_column,
                        release_date_column,
                    ]].copy()
                    distribution_preview_df = distribution_preview_df.rename(columns={
                        amount_column: "Distribution Amount",
                        release_date_column: "Release Date",
                    })

                distribution_preview_df = distribution_preview_df.dropna(subset=["Limited Partner"])
                distribution_preview_df["Limited Partner"] = (
                    distribution_preview_df["Limited Partner"].astype(str).str.strip()
                )
                distribution_preview_df = distribution_preview_df[
                    distribution_preview_df["Limited Partner"] != ""
                ]

                available_distribution_investors = (
                    distribution_preview_df["Limited Partner"]
                    .dropna()
                    .astype(str)
                    .sort_values()
                    .unique()
                    .tolist()
                )
                selected_distribution_investors = st.multiselect(
                    "Investors to include",
                    available_distribution_investors,
                    default=available_distribution_investors,
                    key="selected_distribution_investors",
                )
                distribution_preview_df = distribution_preview_df[
                    distribution_preview_df["Limited Partner"].isin(selected_distribution_investors)
                ]

                st.subheader("Selected Distribution Rows")
                st.write(distribution_preview_df)
                st.success(f"Selected {len(distribution_preview_df)} distribution notice row(s).")

                st.subheader("Distribution Pre-Generation Check")

                distribution_to_investor_name_map = {
                    str(value).strip(): str(key).strip()
                    for key, value in manual_drawdown_name_map.items()
                }

                check_rows = []
                critical_issue_count = 0

                for _, distribution_row in distribution_preview_df.iterrows():
                    distribution_name = str(distribution_row["Limited Partner"]).strip()
                    investor_lookup_name = distribution_to_investor_name_map.get(
                        distribution_name,
                        distribution_name,
                    )
                    investor_candidates = list(dict.fromkeys([
                        investor_lookup_name,
                        distribution_name,
                    ]))

                    investor_match = investor_df[
                        investor_df["Limited Partner"].astype(str).str.strip().isin(investor_candidates)
                    ]

                    contact_candidates = list(dict.fromkeys([
                        manual_contact_name_map.get(investor_lookup_name, investor_lookup_name),
                        manual_contact_name_map.get(distribution_name, distribution_name),
                        distribution_name,
                    ]))
                    contact_match = contact_df[
                        contact_df["Member Name"].astype(str).str.strip().isin(contact_candidates)
                    ]

                    investor_row = investor_match.iloc[0] if len(investor_match) else {}
                    contact_row = contact_match.iloc[0] if len(contact_match) else {}

                    critical_issues = []
                    review_items = []

                    if len(investor_match) == 0:
                        critical_issues.append("Investor master match")
                    if len(contact_match) == 0:
                        critical_issues.append("Contact list match")
                    if is_missing(distribution_row.get("Distribution Amount")):
                        critical_issues.append("Distribution Amount")
                    if is_missing(investor_row.get("Investor_ID", "")):
                        critical_issues.append("Investor ID")
                    if is_missing(contact_row.get("Call Notice Contact Name", "")):
                        critical_issues.append("Contact Name")
                    if is_missing(contact_row.get("Postal Address", "")):
                        critical_issues.append("Postal Address")
                    if is_missing(contact_row.get("Call Notice Contact Email", "")):
                        review_items.append("Email")
                    if is_missing(investor_row.get("Payment_Reference", "")):
                        review_items.append("Payment Reference")

                    if critical_issues:
                        critical_issue_count += 1

                    check_rows.append({
                        "Distribution Name": distribution_name,
                        "Investor Master Name": investor_match.iloc[0]["Limited Partner"] if len(investor_match) else "",
                        "Contact Lookup Name": contact_match.iloc[0]["Member Name"] if len(contact_match) else "",
                        "Critical Issues": ", ".join(critical_issues),
                        "Review Items": ", ".join(review_items),
                        "Status": "Ready" if not critical_issues else "Needs Fix",
                    })

                distribution_check_df = pd.DataFrame(check_rows)

                if critical_issue_count:
                    st.error(f"{critical_issue_count} selected distribution row(s) need fixes before generation.")
                    issue_columns = [
                        "Distribution Name",
                        "Critical Issues",
                        "Review Items",
                        "Status",
                    ]
                    st.write(
                        distribution_check_df[
                            distribution_check_df["Status"] == "Needs Fix"
                        ][issue_columns]
                    )
                else:
                    st.success("All selected distribution rows are ready for document generation.")

                st.write(distribution_check_df[[
                    "Distribution Name",
                    "Investor Master Name",
                    "Contact Lookup Name",
                    "Status",
                ]])

                with st.expander("Full distribution matching audit"):
                    st.write(distribution_check_df)

            if uploaded_distribution_template is None:
                st.info("Upload the distribution notice template before we add document generation.")
            if uploaded_distribution_sample is None:
                st.info("Upload one recently generated notice so we can match the exact wording and layout.")


with monthly_reporting_tab:
    render_monthly_reporting_page()
