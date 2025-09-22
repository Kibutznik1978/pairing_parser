#!/usr/bin/env python3
import re
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image as PILImage
import PyPDF2

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet

PAIR_RE = re.compile(r'\((\d{1,2})\)\s*(\d{1,2}):(\d{2})')
WIN_START = 2*60 + 30
WIN_END   = 5*60

def _to_min(hr: int, mn: int) -> int:
    return (hr % 24) * 60 + mn

def _in_window(hr: int, mn: int) -> bool:
    t = _to_min(hr, mn)
    return WIN_START <= t <= WIN_END

def _crosses_window(dep_h: int, dep_m: int, arr_h: int, arr_m: int) -> bool:
    d = _to_min(dep_h, dep_m)
    a = _to_min(arr_h, arr_m)
    if a < d:
        a += 24 * 60
    return (d <= WIN_START <= a) or (d <= WIN_END <= a)

def _extract_text(pdf_path: Path) -> str:
    reader = PyPDF2.PdfReader(str(pdf_path))
    return "\n".join((p.extract_text() or "") for p in reader.pages)

def _split_trips(text: str) -> List[Tuple[int, str]]:
    import re
    parts = re.split(r'(?:^|\n)Trip Id:\s*(\d+)\s*\n', text)
    trips = []
    for i in range(1, len(parts), 2):
        tid = int(parts[i])
        content = parts[i+1]
        trips.append((tid, content))
    return trips

def _split_duty_blocks(trip_text: str) -> List[str]:
    import re
    parts = re.split(r'\bBriefing\b', trip_text)
    return parts[1:] if len(parts) > 1 else []

def _duty_is_edw(block_text: str) -> bool:
    for line in block_text.splitlines():
        pairs = PAIR_RE.findall(line)
        if len(pairs) >= 2:
            for i in range(len(pairs) - 1):
                (lh1, zh1, zm1) = pairs[i]
                (lh2, zh2, zm2) = pairs[i + 1]
                try:
                    lh1, zm1 = int(lh1), int(zm1)
                    lh2, zm2 = int(lh2), int(zm2)
                except ValueError:
                    continue
                if _in_window(lh1, zm1) or _in_window(lh2, zm2) or _crosses_window(lh1, zm1, lh2, zm2):
                    return True
        elif len(pairs) == 1:
            lh, zh, zm = pairs[0]
            try:
                if _in_window(int(lh), int(zm)):
                    return True
            except ValueError:
                pass
    return False

def analyze_pdf(pdf_path: Path) -> Dict[str, pd.DataFrame]:
    text = _extract_text(pdf_path)
    trips = _split_trips(text)

    records = []
    for tid, content in trips:
        blocks = _split_duty_blocks(content)
        flags = [_duty_is_edw(b) for b in blocks] if blocks else []
        records.append({
            "TripId": tid,
            "DutyDays": len(blocks),
            "EDW_AnyDutyDay": any(flags) if flags else False,
            "EDW_Days_Count": sum(flags) if flags else 0
        })
    trip_flags = pd.DataFrame(records).sort_values("TripId").reset_index(drop=True)

    length_counts = trip_flags["DutyDays"].value_counts().sort_index()
    len_summary = pd.DataFrame({"DutyDays": length_counts.index.astype(int), "Trips": length_counts.values})
    len_summary["Percent"] = (len_summary["Trips"] / len(trip_flags) * 100).round(1)

    total = len(trip_flags)
    edw_trips = int(trip_flags["EDW_AnyDutyDay"].sum())
    edw_pct = round(edw_trips / total * 100, 1) if total else 0.0
    day_trips = total - edw_trips
    day_pct = round(100 - edw_pct, 1)
    edw_summary = pd.DataFrame({
        "Category": ["Day", "EDW (any duty day touches 02:30–05:00)"],
        "Trips": [day_trips, edw_trips],
        "Percent": [day_pct, edw_pct]
    })

    total_days = int(trip_flags["DutyDays"].sum())
    total_edw_days = int(trip_flags["EDW_Days_Count"].sum())
    trip_weighted = round(trip_flags["EDW_AnyDutyDay"].mean() * 100, 1)
    length_weighted_trip = round(trip_flags.loc[trip_flags["EDW_AnyDutyDay"], "DutyDays"].sum() / total_days * 100, 1) if total_days else 0.0
    dutyday_weighted = round(total_edw_days / total_days * 100, 1) if total_days else 0.0

    by_length = trip_flags.groupby("DutyDays").agg(
        Trips=("TripId","count"),
        EDW_Trips=("EDW_AnyDutyDay","sum"),
        EDW_Days=("EDW_Days_Count","sum"),
    ).reset_index()
    by_length["EDW_TripPct"] = (by_length["EDW_Trips"] / by_length["Trips"] * 100).round(1)
    by_length["Total_DutyDays"] = by_length["DutyDays"] * by_length["Trips"]
    by_length["EDW_DutyDayPct"] = (by_length["EDW_Days"] / by_length["Total_DutyDays"] * 100).round(1)

    weighting_summary = pd.DataFrame([{
        "Trip-weighted EDW trip %": f"{trip_weighted}%",
        "Length-weighted EDW trip % (by duty days)": f"{length_weighted_trip}%",
        "Duty-day-weighted EDW day %": f"{dutyday_weighted}%",
    }])

    return {
        "trip_flags": trip_flags,
        "len_summary": len_summary,
        "edw_summary": edw_summary,
        "by_length": by_length,
        "weighting_summary": weighting_summary,
    }

def _save_chart_bar(xlabels, values, title, xlabel, ylabel, out_path: Path):
    import matplotlib.pyplot as plt
    plt.figure()
    plt.bar(xlabels, values)
    plt.title(title)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.tight_layout()
    plt.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close()

def _flatten_to_jpeg(src_path: Path) -> Path:
    dst = src_path.with_suffix(".jpg")
    PILImage.open(src_path).convert("RGB").save(dst, format="JPEG", quality=95)
    return dst

def build_report(outputs: Dict[str, pd.DataFrame], output_dir: Path, domicile: str, aircraft: str, bid_period: str) -> Path:
    len_summary = outputs["len_summary"].sort_values("DutyDays")
    edw_summary = outputs["edw_summary"]
    ws = outputs["weighting_summary"].iloc[0].to_dict()

    chart_trips = output_dir / "trip_length_chart.png"
    chart_pct = output_dir / "trip_length_percent_chart.png"
    chart_edw = output_dir / "edw_vs_day_percent_chart.png"
    chart_three_way = output_dir / "edw_weighting_three_way.png"

    _save_chart_bar(len_summary["DutyDays"].astype(str), len_summary["Trips"],
                    f"Trip Count by Duty Days ({domicile} {aircraft}, Bid {bid_period})",
                    "Duty Days", "Number of Trips", chart_trips)
    _save_chart_bar(len_summary["DutyDays"].astype(str), len_summary["Percent"],
                    f"Percent of Trips by Duty Days ({domicile} {aircraft}, Bid {bid_period})",
                    "Duty Days", "Percent of Trips", chart_pct)
    _save_chart_bar(edw_summary["Category"], edw_summary["Percent"],
                    f"EDW vs Day – Percent of Trips ({domicile} {aircraft}, Bid {bid_period})",
                    "Category", "Percent", chart_edw)
    _save_chart_bar(["Trip-weighted", "Length-weighted (trip)", "Duty-day-weighted (day)"],
                    [float(ws["Trip-weighted EDW trip %"].strip('%')),
                     float(ws["Length-weighted EDW trip % (by duty days)"].strip('%')),
                     float(ws["Duty-day-weighted EDW day %"].strip('%'))],
                    "EDW Share under Different Weightings", "Weighting Method", "Percent EDW", chart_three_way)

    charts_jpg = [_flatten_to_jpeg(p) for p in [chart_trips, chart_pct, chart_edw, chart_three_way]]

    report_pdf = output_dir / f"{domicile}_{aircraft}_Bid{bid_period}_Report_WITH_LengthWeighted.pdf"
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(report_pdf), pagesize=letter, leftMargin=0.7*inch, rightMargin=0.7*inch, topMargin=0.7*inch, bottomMargin=0.7*inch)
    story = []
    max_width = 7.0*inch

    story.append(Paragraph(f"{domicile} {aircraft} – Bid {bid_period}<br/>Trip Length Breakdown", styles["Title"]))
    story.append(Spacer(1, 0.2*inch))
    tbl1_data = [["Duty Days", "Trips", "Percent"]] + [[int(r.DutyDays), int(r.Trips), f"{r.Percent:.1f}%"] for _, r in len_summary.iterrows()]
    tbl1 = Table(tbl1_data, hAlign="LEFT")
    tbl1.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.lightgrey), ("ALIGN", (0,0), (-1,-1), "CENTER"),
                              ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
                              ("BOTTOMPADDING", (0,0), (-1,0), 8)]))
    story.append(tbl1); story.append(Spacer(1, 0.2*inch))
    for p in charts_jpg[:2]:
        img = Image(str(p))
        if img.drawWidth > max_width:
            s = max_width / img.drawWidth; img.drawWidth *= s; img.drawHeight *= s
        story.append(img); story.append(Spacer(1, 0.15*inch))
    story.append(PageBreak())

    story.append(Paragraph(f"{domicile} {aircraft} – Bid {bid_period}<br/>EDW vs Day (Trip-level)", styles["Title"]))
    story.append(Spacer(1, 0.2*inch))
    es = outputs["edw_summary"]
    tbl2_data = [["Category", "Trips", "Percent"]] + [[r.Category, int(r.Trips), f"{r.Percent:.1f}%"] for _, r in es.iterrows()]
    tbl2 = Table(tbl2_data, hAlign="LEFT")
    tbl2.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.lightgrey), ("ALIGN", (0,0), (-1,-1), "CENTER"),
                              ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
                              ("BOTTOMPADDING", (0,0), (-1,0), 8)]))
    story.append(tbl2); story.append(Spacer(1, 0.2*inch))
    img3 = Image(str(charts_jpg[2]))
    if img3.drawWidth > max_width:
        s = max_width / img3.drawWidth; img3.drawWidth *= s; img3.drawHeight *= s
    story.append(img3); story.append(PageBreak())

    story.append(Paragraph(f"{domicile} {aircraft} – Bid {bid_period}<br/>Length‑Weighted EDW Trip Rate", styles["Title"]))
    story.append(Spacer(1, 0.2*inch))
    explain = ("Because many single‑day day trips exist while EDW trips tend to be longer, a simple trip count understates EDW prevalence. "
               "We weight each trip by its length (duty days) while keeping the rule that a trip is EDW if any duty day touches 02:30–05:00 local. "
               "<b>Local time source:</b> the local <i>hour</i> is the number in parentheses, and the <i>minutes</i> come from the following Zulu time. "
               "Example: “(04) 11:22” means <b>local 04:22</b>. EDW window is inclusive of 02:30 and 05:00.")
    story.append(Paragraph(explain, styles["BodyText"])); story.append(Spacer(1, 0.2*inch))
    ws = outputs["weighting_summary"].iloc[0].to_dict()
    t3 = [["Metric", "Value"],
          ["Trip‑weighted EDW trip %", ws["Trip-weighted EDW trip %"]],
          ["Length‑weighted EDW trip % (by duty days)", ws["Length-weighted EDW trip % (by duty days)"]],
          ["Duty‑day‑weighted EDW day %", ws["Duty-day-weighted EDW day %"]]]
    tbl3 = Table(t3, hAlign="LEFT")
    tbl3.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.lightgrey), ("ALIGN", (0,0), (-1,-1), "CENTER"),
                              ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
                              ("BOTTOMPADDING", (0,0), (-1,0), 8)]))
    story.append(tbl3); story.append(Spacer(1, 0.25*inch))
    img4 = Image(str(charts_jpg[-1]))
    if img4.drawWidth > max_width:
        s = max_width / img4.drawWidth; img4.drawWidth *= s; img4.drawHeight *= s
    story.append(img4)

    doc.build(story)
    return report_pdf

def _save_csvs(outputs: Dict[str, pd.DataFrame], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    outputs["trip_flags"].to_csv(output_dir / "trip_level_edw_flags.csv", index=False)
    outputs["len_summary"].to_csv(output_dir / "trip_length_summary.csv", index=False)
    outputs["edw_summary"].to_csv(output_dir / "edw_vs_day_summary.csv", index=False)
    outputs["by_length"].to_csv(output_dir / "edw_by_length.csv", index=False)
    outputs["weighting_summary"].to_csv(output_dir / "edw_weighting_summary.csv", index=False)
    edw_list = outputs["trip_flags"][outputs["trip_flags"]["EDW_AnyDutyDay"]][["TripId","DutyDays","EDW_Days_Count"]].sort_values("TripId")
    edw_list.to_csv(output_dir / "edw_trip_ids.csv", index=False)

def _save_excel(outputs: Dict[str, pd.DataFrame], output_path: Path) -> None:
    with pd.ExcelWriter(output_path) as writer:
        outputs["trip_flags"].to_excel(writer, sheet_name="trip_flags", index=False)
        outputs["len_summary"].to_excel(writer, sheet_name="trip_length_summary", index=False)
        outputs["edw_summary"].to_excel(writer, sheet_name="edw_vs_day_summary", index=False)
        outputs["by_length"].to_excel(writer, sheet_name="edw_by_length", index=False)
        outputs["weighting_summary"].to_excel(writer, sheet_name="edw_weighting_summary", index=False)

def run_edw_report(pdf_path: Path, output_dir: Path, domicile: str="DOM", aircraft: str="AC", bid_period: str="0000"):
    outputs = analyze_pdf(pdf_path)
    _save_csvs(outputs, output_dir)
    report_pdf = build_report(outputs, output_dir, domicile, aircraft, bid_period)
    excel_path = output_dir / f"{domicile}_{aircraft}_Bid{bid_period}_EDW_Report_Data.xlsx"
    _save_excel(outputs, excel_path)
    return {"report_pdf": report_pdf, "excel": excel_path}
