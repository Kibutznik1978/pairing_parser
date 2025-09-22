import re
import unicodedata
from pathlib import Path
from io import BytesIO
import pandas as pd
import matplotlib.pyplot as plt
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
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u00A0", " ")
    text = re.sub(r"[■•▪●]", "-", text)
    return text

# -------------------------------------------------------------------
# PDF Parsing & EDW Logic
# -------------------------------------------------------------------
def parse_pairings(pdf_path: Path):
    reader = PdfReader(str(pdf_path))
    all_text = ""
    for page in reader.pages:
        all_text += page.extract_text() + "\n"

    # Each trip starts with "Trip" or "ID" depending on format
    trips = []
    current_trip = []
    for line in all_text.splitlines():
        if re.match(r"^\s*Trip\s+\d+", line) or re.match(r"^\s*\d+\s+\(", line):
            if current_trip:
                trips.append(current_trip)
                current_trip = []
        current_trip.append(line)
    if current_trip:
        trips.append(current_trip)

    return trips


def extract_local_times(trip_lines):
    """Return list of local times (HH:MM) for a trip."""
    times = []
    pattern = re.compile(r"\((\d{1,2})\)(\d{2}):(\d{2})")
    for line in trip_lines:
        for match in pattern.finditer(line):
            local_hour = int(match.group(1))
            minute = int(match.group(3))
            times.append(f"{local_hour:02d}:{minute:02d}")
    return times


def is_edw_trip(trip_lines):
    """Flag trip as EDW if any local time between 02:30 and 05:00 inclusive."""
    times = extract_local_times(trip_lines)
    for t in times:
        hh, mm = map(int, t.split(":"))
        if (hh == 2 and mm >= 30) or (hh in [3, 4]) or (hh == 5 and mm == 0):
            return True
    return False


# -------------------------------------------------------------------
# PDF Utilities
# -------------------------------------------------------------------
def _save_pdf(title, tables, charts, output_path: Path):
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(clean_text(title), styles["Title"]))
    story.append(Spacer(1, 12))

    for caption, df in tables.items():
        story.append(Paragraph(clean_text(caption), styles["Heading2"]))
        story.append(Spacer(1, 6))

        data = [list(df.columns)] + df.values.tolist()
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
    trips = parse_pairings(pdf_path)

    trip_records = []
    for i, trip_lines in enumerate(trips, start=1):
        edw_flag = is_edw_trip(trip_lines)
        trip_length = sum("DUTY" in l or "LEG" in l for l in trip_lines)  # crude estimate
        trip_records.append({
            "Trip ID": i,
            "Length (days)": trip_length if trip_length > 0 else 1,
            "EDW": edw_flag,
        })

    df_trips = pd.DataFrame(trip_records)

    # Summaries
    total_trips = len(df_trips)
    edw_trips = df_trips["EDW"].sum()
    pct_edw = edw_trips / total_trips * 100

    trip_summary = pd.DataFrame({
        "Metric": ["Total Trips", "EDW Trips", "Day Trips", "Pct EDW"],
        "Value": [total_trips, edw_trips, total_trips - edw_trips, f"{pct_edw:.1f}%"],
    })

    # Weighted summary (trip-weighted vs length-weighted)
    trip_weighted = pct_edw
    length_weighted = df_trips.loc[df_trips["EDW"], "Length (days)"].sum() / df_trips["Length (days)"].sum() * 100

    weighted_summary = pd.DataFrame({
        "Metric": ["Trip-weighted EDW trip %", "Length-weighted EDW trip %"],
        "Value": [f"{trip_weighted:.1f}%", f"{length_weighted:.1f}%"],
    })

    # Save Excel
    excel_path = output_dir / f"{domicile}_{aircraft}_Bid{bid_period}_EDW_Report_Data.xlsx"
    _save_excel({
        "Trip Summary": df_trips,
        "Summary Metrics": trip_summary,
        "Weighted Summary": weighted_summary,
    }, excel_path)

    # Save PDF
    pdf_report_path = output_dir / f"{domicile}_{aircraft}_Bid{bid_period}_EDW_Report.pdf"

    fig, ax = plt.subplots()
    df_trips["EDW"].value_counts().rename({True: "EDW", False: "Day"}).plot(kind="bar", ax=ax)
    ax.set_title("Trips by Type")

    _save_pdf(
        f"{domicile} {aircraft} – Bid {bid_period}",
        {
            "Summary Metrics": trip_summary,
            "Weighted EDW Summary": weighted_summary,
        },
        {"Trips by Type": fig},
        pdf_report_path,
    )

    return {
        "excel": excel_path,
        "report_pdf": pdf_report_path,
        "trip_summary": trip_summary,
        "weighted_summary": weighted_summary,
        "df_trips": df_trips,
    }

