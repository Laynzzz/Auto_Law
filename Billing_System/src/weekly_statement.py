"""Weekly Statement of Account generator.

Builds a summary document per firm for a given business week (Mon-Fri).
No new invoice numbers are assigned — this is a recap only.

Output structure:
    invoice/{FirmName}/{YYYY}/{Mon}/Week of MM-DD-YYYY/Week of MM-DD-YYYY.docx
    invoice/{FirmName}/{YYYY}/{Mon}/Week of MM-DD-YYYY/Week of MM-DD-YYYY.pdf
"""

from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx2pdf import convert

from src.config import get_firm, load_config
from src.dataset import PROJECT_ROOT, query_by_date_range, week_range, _to_date
from src.doc_generator import _format_date_display, _ordinal


# ── Build document ───────────────────────────────────────────────────

TABLE_COLUMNS = ["Date", "Invoice #", "Index #", "Case Caption", "Court", "Amount"]
COL_WIDTHS = [1.0, 1.0, 1.1, 1.8, 1.3, 0.8]  # inches, total ~7.0


def _build_statement(
    firm_name: str,
    monday: date,
    friday: date,
    cases: list[dict],
    output_docx: Path,
) -> Path:
    """Create a weekly statement .docx and save to output_docx."""
    doc = Document()

    # ── Page margins
    for section in doc.sections:
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

    # ── Header
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = h.add_run("PICERNO & ASSOCIATES, PLLC")
    run.bold = True
    run.font.size = Pt(14)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = sub.add_run("Weekly Statement of Account")
    run.font.size = Pt(12)

    # ── Firm & date range
    doc.add_paragraph()
    info = doc.add_paragraph()
    info.add_run("Firm: ").bold = True
    info.add_run(firm_name)

    period = doc.add_paragraph()
    period.add_run("Period: ").bold = True
    period.add_run(
        f"{_format_date_display(monday.isoformat())} - "
        f"{_format_date_display(friday.isoformat())}"
    )

    # ── Disclaimer
    doc.add_paragraph()
    disc = doc.add_paragraph()
    disc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = disc.add_run(
        "This statement summarizes invoices previously sent. "
        "No new charges are added."
    )
    run.italic = True
    run.font.size = Pt(9)

    doc.add_paragraph()

    # ── Table
    table = doc.add_table(rows=1, cols=len(TABLE_COLUMNS))
    table.style = "Light Grid Accent 1"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    for i, name in enumerate(TABLE_COLUMNS):
        cell = table.rows[0].cells[i]
        cell.text = name
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(9)

    # Column widths
    for i, w in enumerate(COL_WIDTHS):
        table.columns[i].width = Inches(w)

    # Data rows
    total = 0.0
    for case in cases:
        row = table.add_row()
        d = _to_date(case.get("appearance_date"))
        date_str = d.strftime("%m/%d/%Y") if d else ""
        amt = float(case.get("charge_amount") or 0)
        total += amt

        values = [
            date_str,
            str(case.get("invoice_number") or ""),
            str(case.get("index_number") or ""),
            str(case.get("case_caption") or ""),
            str(case.get("court") or ""),
            f"${amt:,.2f}",
        ]
        for i, val in enumerate(values):
            cell = row.cells[i]
            cell.text = val
            for p in cell.paragraphs:
                p.alignment = (
                    WD_ALIGN_PARAGRAPH.RIGHT if i == len(values) - 1
                    else WD_ALIGN_PARAGRAPH.LEFT
                )
                for r in p.runs:
                    r.font.size = Pt(9)

    # Total row
    total_row = table.add_row()
    for i in range(len(TABLE_COLUMNS)):
        cell = total_row.cells[i]
        if i == len(TABLE_COLUMNS) - 2:
            cell.text = "Total:"
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for r in p.runs:
                    r.bold = True
                    r.font.size = Pt(9)
        elif i == len(TABLE_COLUMNS) - 1:
            cell.text = f"${total:,.2f}"
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for r in p.runs:
                    r.bold = True
                    r.font.size = Pt(9)

    # ── Count summary
    doc.add_paragraph()
    summary = doc.add_paragraph()
    summary.add_run(f"Total cases: {len(cases)}").font.size = Pt(9)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))
    return output_docx


# ── Full pipeline ────────────────────────────────────────────────────

def generate_weekly_statement(
    firm_name: str,
    week_of: date,
    config: dict | None = None,
    keep_docx: bool = False,
) -> Path:
    """Generate a weekly statement PDF for a firm.

    week_of: any date within the desired week (Mon-Fri range is computed).
    Returns the path to the generated PDF.
    """
    if config is None:
        config = load_config()

    monday, friday = week_range(week_of)
    cases = query_by_date_range(firm_name, monday, friday)

    # Output paths — flat in the week folder (no subfolders)
    date_prefix = monday.strftime("%m-%d-%Y")
    week_folder = f"Week of {date_prefix}"

    base_dir = (
        PROJECT_ROOT / "invoice" / firm_name
        / str(monday.year) / monday.strftime("%b") / week_folder
    )
    docx_out = base_dir / f"{week_folder}.docx"
    pdf_out = base_dir / f"{week_folder}.pdf"

    # Build document
    _build_statement(firm_name, monday, friday, cases, docx_out)

    # Convert to PDF
    pdf_out.parent.mkdir(parents=True, exist_ok=True)
    convert(str(docx_out), str(pdf_out))

    # Clean up intermediate .docx (don't remove week folder — it has other files)
    if not keep_docx and docx_out.exists():
        docx_out.unlink()

    return pdf_out
