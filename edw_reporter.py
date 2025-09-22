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
    PageBreak,
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
    times = []
    pattern = re.compile(r"\((\d{1,2})\)(\d{2}):(\d{2})")
    for match in pattern.finditer(trip_text):
        local_hour = int(match.group(1))
        minute = int(match.group(3))
        times.append(f"{local_hour:02d}:{minute:02d}")
    return times


def is_edw_trip(trip_text):
    times = extract_local_times(trip_text)
    for t in times:
        hh, mm = map(int, t.split(":"))
        if (hh == 2 and mm >= 30) or (hh in [3, 4]) or (hh == 5 and mm == 0):
            return True
    return False


def parse_tafb(trip_text):
    m = re.search(r"TAFB:\s*(\d+)h(\d+)", trip_text)
    if not m:
        return 0.0
    hours = int(m.group(1))
    mins = int(m.group(2))
    return hours + mins / 60.0


def parse_duty_days(trip_text):
    duty_blocks = re.findall(r"(?i)Duty\s+\d+h\d+", trip_text)
    return len(duty_blocks)


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

    # Duty Day distribution (exclude 0s)
    duty_dist = df_trips.groupby("Duty Days")["Trip ID"].count().reset_index(name="Trips")
    duty_dist = duty_dist[duty_dist["Duty Days"] > 0]
    duty_dist["Percent"] = (duty_dist["Trips"] / duty_dist["Trips"].sum() * 100).round(1)

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

    trip_summary = pd.DataFrame({
        "Metric": ["Total Trips", "EDW Trips", "Day Trips", "Pct EDW"],
        "Value": [total_trips, edw_trips, total_trips - edw_trips, f"{trip_weighted:.1f}%"],
    })

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

    # Excel export
    excel_path = output_dir / f"{domicile}_{aircraft}_Bid{bid_period}_EDW_Report_Data.xlsx"
    _save_excel({
        "Trip Records": df_trips,
        "Duty Distribution": duty_dist,
        "Trip Summary": trip_summary,
        "Weighted Summary": weighted_summary,
    }, excel_path)

    # -------------------- Charts --------------------
    # Duty Day Count (bar)
    fig_duty_count, ax1 = plt.subplots()
    ax1.bar(duty_dist["Duty Days"], duty_dist["Trips"])
    for i, v in enumerate(duty_dist["Trips"]):
        ax1.text(duty_dist["Duty Days"].iloc[i], v, str(v), ha="center", va="bottom")
    ax1.set_title("Trips by Duty Day Count")
    ax1.set_xlabel("Duty Days")
    ax1.set_ylabel("Trips")

    # Duty Day Percent (bar)
    fig_duty_percent, ax2 = plt.subplots()
    ax2.bar(duty_dist["Duty Days"], duty_dist["Percent"])
    for i, v in enumerate(duty_dist["Percent"]):
        ax2.text(duty_dist["Duty Days"].iloc[i], v, f"{v:.1f}%", ha="center", va="bottom")
    ax2.set_title("Percentage of Trips by Duty Days")
    ax2.set_xlabel("Duty Days")
    ax2.set_ylabel("Percent")

    # Weighted EDW % (bar)
    fig_edw_bar, ax3 = plt.subplots()
    edw_metrics = ["Pairing %", "Trip-weighted", "TAFB-weighted", "Duty-day-weighted"]
    edw_values = [trip_weighted, trip_weighted, tafb_weighted, dutyday_weighted]
    ax3.bar(edw_metrics, edw_values)
    for i, v in enumerate(edw_values):
        ax3.text(i, v, f"{v:.1f}%", ha="center", va="bottom")
    ax3.set_ylim(0, 100)
    ax3.set_title("EDW Percentages by Method")
    ax3.set_ylabel("Percent")

    # EDW vs Day Trips (pie)
    fig_edw_vs_day, ax4 = plt.subplots()
    ax4.pie([edw_trips, total_trips - edw_trips],
            labels=["EDW Trips", "Day Trips"], autopct="%1.1f%%")
    ax4.set_title("EDW vs Day Trips")

    # Weighted percentages pies
    fig_trip_weight, ax5 = plt.subplots()
    ax5.pie([trip_weighted, 100 - trip_weighted], labels=["EDW", "Day"], autopct="%1.1f%%")
    ax5.set_title("Trip-weighted EDW %")

    fig_tafb_weight, ax6 = plt.subplots()
    ax6.pie([tafb_weighted, 100 - tafb_weighted], labels=["EDW", "Day"], autopct="%1.1f%%")
    ax6.set_title("TAFB-weighted EDW %")

    fig_dutyday_weight, ax7 = plt.subplots()
    ax7.pie([dutyday_weighted, 100 - dutyday_weighted], labels=["EDW", "Day"], autopct="%1.1f%%")
    ax7.set_title("Duty-day-weighted EDW %")

    # -------------------- PDF Build --------------------
    pdf_report_path = output_dir / f"{domicile}_{aircraft}_Bid{bid_period}_EDW_Report.pdf"
    styles = getSampleStyleSheet()
    story = []

    # Page 1 – Duty breakdown
    story.append(Paragraph(f"{domicile} {aircraft} – Bid {bid_period} Trip Length Breakdown", styles["Title"]))
    story.append(Spacer(1, 12))
    data = [list(duty_dist.columns)] + duty_dist.values.tolist()
    data = [[clean_text(cell) for cell in row] for row in data]
    t = Table(data)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    story.append(t)
    story.append(Spacer(1, 12))
    for fig in [fig_duty_count, fig_duty_percent]:
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight")
        buf.seek(0)
        img = PILImage.open(buf)
        img_path = output_dir / f"chart_{hash(fig)}.png"
        img.save(img_path)
        story.append(Image(str(img_path), width=400, height=300))
        story.append(Spacer(1, 12))
    story.append(PageBreak())

    # Page 2 – EDW breakdown
    story.append(Paragraph(f"{domicile} {aircraft} – Bid {bid_period} EDW Breakdown", styles["Title"]))
    story.append(Spacer(1, 12))
    for caption, df in {"Trip Summary": trip_summary, "Weighted Summary": weighted_summary}.items():
        story.append(Paragraph(clean_text(caption), styles["Heading2"]))
        data = [list(df.columns)] + df.values.tolist()
        data = [[clean_text(cell) for cell in row] for row in data]
        t = Table(data)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ]))
        story.append(t)
        story.append(Spacer(1, 12))
    buf = BytesIO()
    fig_edw_bar.savefig(buf, format="png", bbox_inches="tight")
    buf.seek(0)
    img = PILImage.open(buf)
    img_path = output_dir / "EDW_Bar.png"
    img.save(img_path)
    story.append(Image(str(img_path), width=400, height=300))
    story.append(PageBreak())

    # Page 3 – Pies
    for fig in [fig_edw_vs_day, fig_trip_weight, fig_tafb_weight, fig_dutyday_weight]:
        buf = BytesIO()
        fig.savefig(buf, format="png", bbox_inches="tight")
        buf.seek(0)
        img = PILImage.open(buf)
        img_path = output_dir / f"pie_{hash(fig)}.png"
        img.save(img_path)
        story.append(Image(str(img_path), width=300, height=300))
        story.append(Spacer(1, 12))

    doc = SimpleDocTemplate(str(pdf_report_path), pagesize=letter)
    doc.build(story)

    return {
        "excel": excel_path,
        "report_pdf": pdf_report_path,
        "df_trips": df_trips,
        "duty_dist": duty_dist,
        "trip_summary": trip_summary,
        "weighted_summary": weighted_summary,
    }



