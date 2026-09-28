from __future__ import annotations

import io
import re
import tempfile
from calendar import month_abbr
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import streamlit as st

from monthly_reporting_automation import PROFILES, find_month_columns, workbook_sheet_name


PROJECT_ROOT = Path(__file__).resolve().parent
QUARTERLY_OUTPUTS_DIR = PROJECT_ROOT / "Outputs" / "Quarterly Reports"

MAX_BULLETS = 4
MAX_BULLET_CHARS = 190
MAX_DESCRIPTION_CHARS = 520
MAX_UPDATES_CHARS = 820

EXCLUDED_LABEL_TOKENS = (
    "source:",
    "note:",
    "check",
    "commentary",
    "what worked",
    "what didn't work",
    "what didnt work",
    "targets for next month",
    "key updates",
    "company description",
    "key financial",
    "standard reporting template",
    "forecast",
)

METRIC_PRIORITIES = (
    ("total revenue", 120),
    ("net revenue", 115),
    ("revenue", 110),
    ("gmv", 105),
    ("gtv", 105),
    ("ntv", 105),
    ("gross profit", 100),
    ("gross margin", 96),
    ("total expenses", 92),
    ("net burn", 90),
    ("ebitda", 88),
    ("net income", 86),
    ("operating profit", 84),
    ("monthly active users", 80),
    ("mau", 80),
    ("bookings", 78),
    ("transactions", 76),
    ("active corporates", 74),
    ("cash", 72),
    ("runway", 70),
    ("subscriptions", 68),
    ("total students", 66),
    ("teacher", 64),
)


@dataclass(frozen=True)
class MetricRow:
    label: str
    values: dict[tuple[int, int], object]


@dataclass(frozen=True)
class CompanyReportData:
    company: str
    profile_name: str
    sheet_name: str | None
    description: str
    key_updates: tuple[str, ...]
    metrics: tuple[MetricRow, ...]
    warnings: tuple[str, ...]


def quarter_months(year: int, quarter_end_month: int) -> tuple[tuple[int, int], ...]:
    quarter_start_month = ((quarter_end_month - 1) // 3) * 3 + 1
    return tuple((year, quarter_start_month + offset) for offset in range(3))


def quarter_label(months: Iterable[tuple[int, int]]) -> str:
    months = tuple(months)
    if not months:
        return "Quarterly Report"
    year, ending_month = months[-1]
    quarter = ((ending_month - 1) // 3) + 1
    return f"Q{quarter} {year}"


def _cell_text(value: object) -> str:
    if not isinstance(value, str):
        return ""
    text = str(value).replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _normalise(value: str) -> str:
    return " ".join(value.lower().replace("&", "and").split())


def _strip_bullet(value: str) -> str:
    text = value.strip()
    text = re.sub(r"^[\u2022*\-]+\s*", "", text)
    return text.strip()


def _first_nonempty_text_in_row(sheet: Worksheet, row: int, max_col: int = 10) -> str:
    for col in range(1, min(sheet.max_column, max_col) + 1):
        text = _cell_text(sheet.cell(row=row, column=col).value)
        if text:
            return text
    return ""


def _find_heading_row(sheet: Worksheet, heading: str, max_row: int = 80) -> tuple[int, int] | None:
    wanted = _normalise(heading)
    for row in range(1, min(sheet.max_row, max_row) + 1):
        for col in range(1, min(sheet.max_column, 10) + 1):
            text = _cell_text(sheet.cell(row=row, column=col).value)
            if wanted in _normalise(text):
                return row, col
    return None


def _extract_description(sheet: Worksheet, company: str) -> str:
    heading = _find_heading_row(sheet, "Company Description", max_row=25)
    if heading is not None:
        row, _col = heading
        candidates = [
            _first_nonempty_text_in_row(sheet, next_row, max_col=8)
            for next_row in range(row + 1, min(row + 6, sheet.max_row) + 1)
        ]
        return max(candidates, key=len, default="").strip()

    company_key = _normalise(company)
    candidates: list[str] = []
    for row in range(1, min(sheet.max_row, 10) + 1):
        text = _first_nonempty_text_in_row(sheet, row, max_col=8)
        if not text:
            continue
        text_key = _normalise(text)
        if text_key == company_key or "key financial" in text_key or "key updates" in text_key:
            continue
        if len(text) >= 24:
            candidates.append(text)
    return max(candidates, key=len, default="").strip()


def _extract_key_updates(sheet: Worksheet) -> tuple[str, ...]:
    heading = _find_heading_row(sheet, "Key Updates", max_row=35)
    if heading is None:
        return ()

    start_row, _col = heading
    updates: list[str] = []
    blank_streak = 0
    for row in range(start_row + 1, min(start_row + 12, sheet.max_row) + 1):
        text = _first_nonempty_text_in_row(sheet, row, max_col=8)
        key = _normalise(text)
        if not text:
            blank_streak += 1
            if blank_streak >= 3 and updates:
                break
            continue
        blank_streak = 0
        if "key financial" in key or "reporting link" in key or "source:" in key:
            break
        updates.append(_strip_bullet(text))
    return tuple(update for update in updates if update)


def _label_for_row(sheet: Worksheet, row: int, first_month_col: int) -> str:
    for col in range(max(1, first_month_col - 6), 0, -1):
        text = _cell_text(sheet.cell(row=row, column=col).value)
        if text:
            return text
    return ""


def _is_metric_value(value: object) -> bool:
    if value is None or isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    return False


def _metric_score(label: str) -> int:
    key = _normalise(label)
    if not key or any(token in key for token in EXCLUDED_LABEL_TOKENS):
        return 0
    for token, score in METRIC_PRIORITIES:
        if token in key:
            return score
    return 25 if len(label) <= 60 else 0


def _extract_metrics(
    sheet: Worksheet,
    months: tuple[tuple[int, int], ...],
) -> tuple[tuple[MetricRow, ...], tuple[str, ...]]:
    warnings: list[str] = []
    month_cols = find_month_columns(sheet, months)
    missing_months = [month for month in months if month not in month_cols]
    if missing_months:
        missing = ", ".join(f"{month_abbr[month]} {str(year)[-2:]}" for year, month in missing_months)
        warnings.append(f"Missing month columns: {missing}.")
    if not month_cols:
        return (), tuple(warnings)

    first_month_col = min(month_cols.values())
    candidates: list[tuple[int, int, MetricRow]] = []
    seen_labels: set[str] = set()
    for row in range(1, min(sheet.max_row, 130) + 1):
        label = _label_for_row(sheet, row, first_month_col)
        score = _metric_score(label)
        if not score:
            continue
        label_key = _normalise(label)
        if label_key in seen_labels:
            continue
        values = {month: sheet.cell(row=row, column=col).value for month, col in month_cols.items()}
        if not any(_is_metric_value(value) for value in values.values()):
            continue
        seen_labels.add(label_key)
        candidates.append((-score, row, MetricRow(label=label, values=values)))

    candidates.sort()
    selected = tuple(metric for _score, _row, metric in candidates[:7])
    if not selected:
        warnings.append("No usable financial or operating metrics found.")
    else:
        for metric in selected:
            missing_values = [month for month in months if not _is_metric_value(metric.values.get(month))]
            if missing_values:
                warnings.append(f"Metric '{metric.label}' has missing values for part of the quarter.")
                break
    return selected, tuple(warnings)


def _text_warnings(description: str, updates: tuple[str, ...]) -> list[str]:
    warnings: list[str] = []
    if not description:
        warnings.append("Missing company description.")
    elif len(description) > MAX_DESCRIPTION_CHARS:
        warnings.append(f"Description may overflow fixed text box ({len(description)} chars).")

    if not updates:
        warnings.append("Missing key updates.")
    if len(updates) > MAX_BULLETS:
        warnings.append(f"Too many key update bullets ({len(updates)}; target {MAX_BULLETS}).")
    if sum(len(update) for update in updates) > MAX_UPDATES_CHARS:
        warnings.append("Key updates may overflow fixed text box.")
    if any(len(update) > MAX_BULLET_CHARS for update in updates):
        warnings.append(f"One or more bullets exceed {MAX_BULLET_CHARS} characters.")
    return warnings


def parse_monthly_reporting_workbook(
    workbook_path: Path,
    months: tuple[tuple[int, int], ...],
) -> tuple[CompanyReportData, ...]:
    workbook = load_workbook(workbook_path, data_only=True)
    companies: list[CompanyReportData] = []
    try:
        for profile_name, profile in PROFILES.items():
            company_warnings: list[str] = []
            sheet_name = workbook_sheet_name(workbook, profile.report_sheet)
            if sheet_name is None:
                companies.append(
                    CompanyReportData(
                        company=profile.display_name,
                        profile_name=profile_name,
                        sheet_name=None,
                        description="",
                        key_updates=(),
                        metrics=(),
                        warnings=(f"Missing workbook sheet '{profile.report_sheet}'.",),
                    )
                )
                continue

            sheet = workbook[sheet_name]
            description = _extract_description(sheet, profile.display_name)
            key_updates = _extract_key_updates(sheet)
            metrics, metric_warnings = _extract_metrics(sheet, months)
            company_warnings.extend(_text_warnings(description, key_updates))
            company_warnings.extend(metric_warnings)
            companies.append(
                CompanyReportData(
                    company=profile.display_name,
                    profile_name=profile_name,
                    sheet_name=sheet_name,
                    description=description,
                    key_updates=key_updates,
                    metrics=metrics,
                    warnings=tuple(company_warnings),
                )
            )
    finally:
        workbook.close()
    return tuple(companies)


def _format_metric_value(label: str, value: object) -> str:
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return "-"

    label_key = _normalise(label)
    if any(token in label_key for token in ("%", "margin", "rate", "yield")) and abs(value) <= 5:
        return f"{value * 100:.1f}%"
    if abs(value) >= 1_000_000:
        return f"{value / 1_000_000:.1f}M"
    if abs(value) >= 1_000:
        return f"{value / 1_000:.1f}K"
    if float(value).is_integer():
        return f"{int(value):,}"
    return f"{value:.1f}"


def _add_textbox(slide, x, y, w, h, text, font_size=12, bold=False, color="1F2937"):
    from pptx.util import Inches, Pt
    from pptx.enum.text import MSO_AUTO_SIZE

    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.NONE
    paragraph = frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = text
    run.font.name = "Aptos"
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = _rgb(color)
    return box


def _rgb(hex_color: str):
    from pptx.dml.color import RGBColor

    value = hex_color.strip().lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def _add_rect(slide, x, y, w, h, fill, line="FFFFFF", transparency=0):
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Inches

    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(fill)
    shape.fill.transparency = transparency
    shape.line.color.rgb = _rgb(line)
    return shape


def _add_metric_table(slide, company: CompanyReportData, months: tuple[tuple[int, int], ...]) -> None:
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt

    rows = max(2, len(company.metrics) + 1)
    cols = len(months) + 1
    table_shape = slide.shapes.add_table(rows, cols, Inches(6.35), Inches(2.0), Inches(6.35), Inches(3.85))
    table = table_shape.table
    widths = [2.75, 1.2, 1.2, 1.2]
    for idx, width in enumerate(widths[:cols]):
        table.columns[idx].width = Inches(width)

    headers = ["Metric", *[f"{month_abbr[month]} {str(year)[-2:]}" for year, month in months]]
    for col, header in enumerate(headers):
        cell = table.cell(0, col)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb("123D40")
        for paragraph in cell.text_frame.paragraphs:
            paragraph.alignment = PP_ALIGN.CENTER if col else PP_ALIGN.LEFT
            for run in paragraph.runs:
                run.font.name = "Aptos"
                run.font.size = Pt(10)
                run.font.bold = True
                run.font.color.rgb = _rgb("FFFFFF")

    for row_idx, metric in enumerate(company.metrics, start=1):
        values = [metric.label, *[_format_metric_value(metric.label, metric.values.get(month)) for month in months]]
        for col, value in enumerate(values):
            cell = table.cell(row_idx, col)
            cell.text = value
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb("F7F9F8") if row_idx % 2 else _rgb("FFFFFF")
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.RIGHT if col else PP_ALIGN.LEFT
                for run in paragraph.runs:
                    run.font.name = "Aptos"
                    run.font.size = Pt(9.2)
                    run.font.color.rgb = _rgb("1F2937")


def _add_bullets(slide, updates: Iterable[str]) -> None:
    from pptx.util import Inches, Pt

    box = slide.shapes.add_textbox(Inches(0.72), Inches(3.7), Inches(5.25), Inches(2.35))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    for idx, update in enumerate(list(updates)[:MAX_BULLETS]):
        paragraph = frame.paragraphs[0] if idx == 0 else frame.add_paragraph()
        paragraph.text = _strip_bullet(update)
        paragraph.level = 0
        paragraph.font.name = "Aptos"
        paragraph.font.size = Pt(12)
        paragraph.space_after = Pt(7)
        paragraph._p.get_or_add_pPr().set("marL", "171450")
        paragraph._p.get_or_add_pPr().set("indent", "-171450")


def build_company_pages_pptx(
    companies: tuple[CompanyReportData, ...],
    review_rows: list[dict[str, object]],
    months: tuple[tuple[int, int], ...],
    output_path: Path,
) -> None:
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
    except ImportError as exc:
        raise RuntimeError("Quarterly Report generation needs python-pptx. Add python-pptx to requirements.txt and install dependencies.") from exc

    data_by_company = {company.company: company for company in companies}
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    report_label = quarter_label(months)

    for review in review_rows:
        if not review.get("Include", True):
            continue
        company = data_by_company.get(str(review["Company"]))
        if company is None:
            continue

        slide = prs.slides.add_slide(blank_layout)
        _add_rect(slide, 0, 0, 13.333, 7.5, "FFFFFF")
        _add_rect(slide, 0, 0, 13.333, 0.24, "123D40", line="123D40")
        _add_rect(slide, 0, 7.2, 13.333, 0.3, "123D40", line="123D40")
        _add_rect(slide, 0.55, 0.62, 0.08, 5.55, "C49A5A", line="C49A5A")

        _add_textbox(slide, 0.72, 0.48, 7.6, 0.45, company.company, font_size=26, bold=True, color="123D40")
        _add_textbox(slide, 10.85, 0.58, 1.75, 0.28, report_label, font_size=11, bold=True, color="5B6770")
        _add_textbox(slide, 0.72, 1.16, 5.2, 0.24, "Company Description", font_size=10.5, bold=True, color="C49A5A")
        _add_textbox(slide, 0.72, 1.48, 5.25, 1.55, str(review.get("Description", "")).strip(), font_size=11.2, color="1F2937")
        _add_textbox(slide, 0.72, 3.36, 5.2, 0.24, "Key Updates", font_size=10.5, bold=True, color="C49A5A")
        updates = str(review.get("Key updates", "")).replace("\r", "\n").splitlines()
        _add_bullets(slide, [update for update in updates if update.strip()])

        _add_textbox(slide, 6.35, 1.16, 4.4, 0.26, "Key Financial & Operating Metrics", font_size=10.5, bold=True, color="C49A5A")
        _add_metric_table(slide, company, months)
        notes = str(review.get("Notes", "")).strip()
        if notes:
            _add_textbox(slide, 6.35, 6.05, 5.95, 0.38, f"Notes: {notes}", font_size=8.8, color="5B6770")
        _add_textbox(slide, 0.72, 6.86, 8.6, 0.18, "Source: Monthly Reporting workbook. Draft for management review.", font_size=7.5, color="FFFFFF")

        for shape in slide.shapes:
            if hasattr(shape, "text_frame"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Aptos"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)


def _review_rows(companies: tuple[CompanyReportData, ...]) -> list[dict[str, object]]:
    return [
        {
            "Include": bool(company.sheet_name),
            "Company": company.company,
            "Description": company.description,
            "Key updates": "\n".join(company.key_updates),
            "Notes": "",
        }
        for company in companies
    ]


def _warning_rows(
    companies: tuple[CompanyReportData, ...],
    edited_rows: list[dict[str, object]] | None = None,
) -> list[dict[str, str]]:
    by_company = {row.get("Company"): row for row in edited_rows or []}
    rows: list[dict[str, str]] = []
    for company in companies:
        warnings = list(company.warnings)
        if company.company in by_company:
            row = by_company[company.company]
            description = str(row.get("Description", "")).strip()
            updates = tuple(
                _strip_bullet(line)
                for line in str(row.get("Key updates", "")).splitlines()
                if line.strip()
            )
            warnings = [warning for warning in warnings if not warning.startswith(("Description", "Missing company", "Missing key", "Too many", "Key updates", "One or more bullets"))]
            warnings.extend(_text_warnings(description, updates))
        for warning in warnings:
            rows.append({"Company": company.company, "Warning": warning})
    return rows


def _latest_quarterly_outputs(limit: int = 8) -> list[dict[str, str]]:
    if not QUARTERLY_OUTPUTS_DIR.exists():
        return []
    files = sorted(QUARTERLY_OUTPUTS_DIR.glob("*.pptx"), key=lambda path: path.stat().st_mtime, reverse=True)
    return [
        {
            "File": path.name,
            "Modified": datetime.fromtimestamp(path.stat().st_mtime).strftime("%b %d, %Y %H:%M"),
            "Folder": str(path.parent),
        }
        for path in files[:limit]
    ]


def render_quarterly_report_page() -> None:
    st.header("Quarterly Report")
    st.caption("Generate editable company-page slides from the completed Monthly Reporting workbook. The Adobe XD PDF is used only as a format reference.")

    first_historical_year = 2021
    year_choices = list(range(first_historical_year, date.today().year + 16))
    quarter_end_options = [("Mar", 3), ("Jun", 6), ("Sep", 9), ("Dec", 12)]

    with st.container(border=True):
        st.subheader("Run setup")
        setup_cols = st.columns([1, 1, 2])
        with setup_cols[0]:
            selected_year = st.selectbox(
                "Year",
                options=year_choices,
                index=year_choices.index(2026) if 2026 in year_choices else len(year_choices) - 1,
                key="quarterly_report_year",
            )
        with setup_cols[1]:
            selected_quarter_label = st.selectbox(
                "Quarter ending",
                options=[label for label, _month in quarter_end_options],
                index=0,
                key="quarterly_report_quarter_end",
            )
        months = quarter_months(selected_year, dict(quarter_end_options)[selected_quarter_label])
        with setup_cols[2]:
            st.metric("Report quarter", quarter_label(months))

        monthly_report_file = st.file_uploader(
            "Completed Monthly Reporting workbook",
            type=["xlsx"],
            key="quarterly_monthly_report_file",
        )

    if monthly_report_file is None:
        st.info("Upload the completed Monthly Reporting workbook to build the review screen.")
        latest_rows = _latest_quarterly_outputs()
        if latest_rows:
            with st.expander("Latest generated quarterly files", expanded=False):
                st.dataframe(latest_rows, hide_index=True, width="stretch")
        return

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        workbook_path = tmp_dir / "monthly_reporting.xlsx"
        workbook_path.write_bytes(monthly_report_file.getbuffer())
        with st.spinner("Reading company pages from the monthly workbook"):
            companies = parse_monthly_reporting_workbook(workbook_path, months)

    parsed_count = sum(1 for company in companies if company.sheet_name)
    warning_rows = _warning_rows(companies)

    with st.container(border=True):
        st.subheader("Workbook scan")
        metric_cols = st.columns(3)
        metric_cols[0].metric("Company sheets found", parsed_count)
        metric_cols[1].metric("Companies in scope", len(companies))
        metric_cols[2].metric("Warnings", len(warning_rows))
        if warning_rows:
            st.dataframe(warning_rows, hide_index=True, width="stretch")
        else:
            st.success("No missing-data or overflow warnings detected.")

    with st.container(border=True):
        st.subheader("Human review")
        st.caption("Edit text only. Metrics are read from the workbook and placed into fixed slide boxes.")
        edited_rows = st.data_editor(
            _review_rows(companies),
            hide_index=True,
            width="stretch",
            disabled=["Company"],
            key=f"quarterly_report_review_{quarter_label(months)}_{monthly_report_file.name}",
            column_config={
                "Include": st.column_config.CheckboxColumn("Include"),
                "Description": st.column_config.TextColumn("Description", width="large"),
                "Key updates": st.column_config.TextColumn("Key updates", width="large"),
                "Notes": st.column_config.TextColumn("Notes", width="medium"),
            },
        )

    edited_warnings = _warning_rows(companies, edited_rows)
    if edited_warnings:
        with st.expander("Warnings after edits", expanded=True):
            st.dataframe(edited_warnings, hide_index=True, width="stretch")

    included_count = sum(1 for row in edited_rows if row.get("Include"))
    generate = st.button(
        "Generate company pages PPTX",
        type="primary",
        disabled=included_count == 0,
        width="stretch",
    )
    if not generate:
        return

    filename = f"Sarmayacar {quarter_label(months)} Quarterly Company Pages.pptx"
    with tempfile.TemporaryDirectory() as tmp:
        output_path = Path(tmp) / filename
        try:
            build_company_pages_pptx(companies, edited_rows, months, output_path)
        except RuntimeError as exc:
            st.error(str(exc))
            return
        pptx_bytes = output_path.read_bytes()

    QUARTERLY_OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    permanent_output_path = QUARTERLY_OUTPUTS_DIR / filename
    permanent_output_path.write_bytes(pptx_bytes)

    with st.container(border=True):
        st.success("Generated! Your editable quarterly company-page deck is ready.")
        st.caption(f"Saved final file: {permanent_output_path}")
        st.download_button(
            "Download Quarterly Report PPTX",
            data=io.BytesIO(pptx_bytes),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            width="stretch",
        )
