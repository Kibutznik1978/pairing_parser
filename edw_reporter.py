import re
import unicodedata
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from PIL import Image as PILImage

from PyPDF2 import PdfReader
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image,
)
from reportlab.lib.styles import getSampleStyleSheet

# -------------------------------------------------------------------
# Text Sanitizer
# -------------------------------------------------------------------
def clean_text(text: str) -> str:
    """
    Normalize text so ReportLab doesn't render black boxes (■).
    - Normalize Unicode (smart quotes → plain quotes, em-dash → hyphen, etc.)
    - Replace non-breaking spaces with regular spaces
    - Replace bullets/black squares with plain dashes
    """
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u00A0", " ")  # non-breaking space → space
    text = re.sub(r"[■•▪●]", "-", text)  # any bullets/boxes → dash
    return text

# -------------------------------------------------------------------
# PDF Utilities
# -------------------------------------------------------------------
def _save_pdf(title, tables, charts, output_path: Path):
    styles = getSampleStyleSheet()
    story = []

    # Title
    story.append(Paragraph(clean_text(title), styles["Title"]))
    story.append(Spacer(1, 12))

    # Add tables
    for caption, df in tables.items():
        story.append(Paragraph(clean_text(caption), styles["Heading2"]))
        story.append(Spacer(1, 6))

        data = [list(df.columns)] + df.values.tolist()
        # clean text inside table
        data = [[clean_text(cell) for cell in row] for row in data]

        t = Table(data)
        t.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ]
            )
        )
        story.append(t)
        story.append(Spacer(1, 12))

    # Add charts
    for caption, fig in charts.items():
        story.append(Paragraph(clean_text(caption), styles["Heading2"]))
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight")
        buf.seek(0)
        img = PILImage.open(buf)
        img_path = output_path.parent / f"{caption}.png"
        img.save(img_path)
        story.append(Image(str(img_path), width=400, height=300))
        story.append(Spacer(1, 12))

    doc = SimpleDocTemplate(str(output_path), pagesize=letter)
    doc.build(story)

# -------------------------------------------------------------------
# Excel Utilities
# -------------------------------------------------------------------
def _save_excel(dfs: dict, output_path: Path):
    with pd.ExcelWriter(output_path) as writer:
        for sheet, df in dfs.items():
            df.to_excel(writer, sheet_name=clean_text(sheet), index=False)

# -------------------------------------------------------------------
# Main Reporting
# -------------------------------------------------------------------
def run_edw_report(pdf_path: Path, output_dir: Path, domicile: str, aircraft: str, bid_period: str):
    """
    Parse a pairings PDF, classify EDW vs Day trips, generate summaries,
    and export Excel + PDF reports.
    """

    # Dummy placeholders for parsed data (replace with your real parsing logic)
    trip_summary = pd.DataFrame({
        "Trip ID": [1, 2, 3],
        "Type": ["Day", "EDW", "EDW"],
        "Length": [1, 4, 6],
    })

    weighted_summary = pd.DataFrame({
        "Metric": ["Trip-weighted EDW trip %", "Length-weighted EDW trip %", "Duty-day-weighted EDW day %"],
        "Value": ["65.6%", "77.4%", "47.0%"],
    })

    # Save Excel
    excel_path = output_dir / f"{domicile}_{aircraft}_Bid{bid_period}_EDW_Report_Data.xlsx"
    _save_excel({
        "Trip Summary": trip_summary,
        "Weighted Summary": weighted_summary,
    }, excel_path)

    # Save PDF
    pdf_report_path = output_dir / f"{domicile}_{aircraft}_Bid{bid_period}_EDW_Report.pdf"

    # Example chart
    fig, ax = plt.subplots()
    trip_summary["Type"].value_counts().plot(kind="bar", ax=ax)
    ax.set_title("Trips by Type")

    _save_pdf(
        f"{domicile} {aircraft} – Bid {bid_period}",
        {"Weighted EDW Summary": weighted_summary},
        {"Trips by Type": fig},
        pdf_report_path,
    )

    return {
        "excel": excel_path,
        "report_pdf": pdf_report_path,
        "trip_summary": trip_summary,
        "weighted_summary": weighted_summary,
    }

