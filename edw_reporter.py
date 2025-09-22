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

    # Split trips (each trip starts with "Trip Id" or "Trip")
    trips = []
    current_trip = []
    for line in all_text.splitlines():
        if re.match(r"^\s*Trip\s*Id", line, re.IGNORECASE):
            if current_trip:
                trips.append("\n".join(current_trip))
                current_trip = []
        current_trip.append(line)
    if current_trip:
        trips.append("\n".join(current_trip))

    return trips


def extract_local_times(trip_text):
    """Return list of local times (HH:MM) for a trip."""
    times = []
    pattern = re.compile(r"\((\d{1,2})\)(\d{2}):(\d{2})")
    for match in pattern.finditer(trip_text):
        local_hour = int(match.group(1))
        minute = int(match.group(3))
        times.append(f"{local_hour:02d}:{minute:02d}")
    return times


def is_edw_trip(trip_text):
    """Flag trip as EDW if any local time between 02:30 and 05:00 inclusive."""
    times = extract_local_times(trip_text)
    for t in times:
        hh, mm = map(int, t.split(":"))
        if (hh == 2 and mm >= 30) or (hh in [3, 4]) or (hh == 5 and mm == 0):
            return True
    return False


def parse_tafb(trip_text):
    """Extract TAFB in hours from Trip Summary."""
    m = re.search(r"TAFB:\s*(\d+)h(\d+)", trip_text)
    if not m:
        return 0.0
    hours = int(m.group(1))
    mins = int(m.group(2))
    return hours + mins / 60.0


def parse_duty_days(trip_text):
    """Count how many Duty blocks appear in Duty Summary."""
    duty_blocks = re.findall(r"(?i)Duty\s+\d+h\d+", trip_text)
    return len(duty_blocks)


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
    for i, trip_text in enumerate(trips, start=1):
        edw_flag = is_edw_trip(trip_text)
        tafb_hours = parse_tafb(trip_text)
        tafb_days = tafb_hours / 24.0 if tafb_hours else 0.0
        duty_days = parse_duty_days(trip_text)

        trip_records.append({
            "Trip ID": i,
            "TAFB Hours": round(tafb_hours, 2),
            "TAFB Days": round(tafb_days, 2),
            "Duty Days": duty_days,
            "EDW": edw_flag,
        })

    df_trips = pd.DataFrame(trip_records)

    # Summaries
    total_trips = len(df_trips)
    edw_trips = df_trips["EDW"].sum()

    trip_weighted = edw_trips / total_trips * 100 if total_trips else 0
    tafb_weighted = (
        df_trips.loc[df_trips["EDW"], "TAFB Hours"].sum()
        / df_trips["TAFB Hours"].sum()
        * 100
        if df_trips["TAFB Hours"].sum() > 0 else 0
    )
    dutyday_weighted = (
        df_trips.loc[df_trips["EDW"], "Duty Days"].sum()
        / df_trips["Duty Days"].sum()
        * 100
        if df_trips["Duty Days"].sum() > 0 else 0
    )

    weighted_summary = pd.DataFrame({
        "Metric": [
            "Trip-weighted EDW trip %",
            "TAFB-weighted EDW trip %",
            "Duty-day-weighted EDW trip %",
        ],
        "Value": [
            f"{trip_weighted:.1f}%",
            f"{tafb_weighted:.1f}%",
            f"{dutyday_weighted:.1f}%",
        ],
    })

    # Save Excel
    excel_path = output_dir / f"{domicile}_{aircraft}_Bid{bid_period}_EDW_Report_Data.xlsx"
    _save_excel({
        "Trip Records": df_trips,
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
            "Weighted Summary": weighted_summary,
        },
        {"Trips by Type": fig},
        pdf_report_path,
    )

    return {
        "excel": excel_path,
        "report_pdf": pdf_report_path,
        "weighted_summary": weighted_summary,
        "df_trips": df_trips,
    }


