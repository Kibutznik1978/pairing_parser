import tempfile
from pathlib import Path
import streamlit as st

# Import your analysis helper
from edw_reporter import run_edw_report

st.set_page_config(page_title="EDW Pairing Analyzer", layout="centered")
st.title("EDW Pairing Analyzer")

st.markdown(
    "Upload a formatted bid-pack PDF. I’ll return an **Excel** workbook and a **3-page PDF** "
    "with the trip-length breakdown, EDW vs Day, and length-weighted explanation."
)

with st.expander("Labels (optional)"):
    dom = st.text_input("Domicile", value="ONT")
    ac  = st.text_input("Aircraft", value="757")
    bid = st.text_input("Bid period", value="2507")

uploaded = st.file_uploader("Pairings PDF", type=["pdf"])

run = st.button("Run Analysis", disabled=(uploaded is None))
if run:
    if uploaded is None:
        st.warning("Please upload a PDF first.")
        st.stop()

    with st.spinner("Crunching your file…"):
        tmpdir = Path(tempfile.mkdtemp())
        pdf_path = tmpdir / uploaded.name
        pdf_path.write_bytes(uploaded.getvalue())

        out_dir = tmpdir / "outputs"
        out_dir.mkdir(exist_ok=True)

        results = run_edw_report(
            pdf_path,
            out_dir,
            domicile=dom.strip() or "DOM",
            aircraft=ac.strip() or "AC",
            bid_period=bid.strip() or "0000"
        )

    st.success("Done! Download your files below:")

    # Excel workbook
    xlsx = out_dir / f"{dom}_{ac}_Bid{bid}_EDW_Report_Data.xlsx"
    st.download_button(
        "⬇️ Download Excel",
        data=xlsx.read_bytes(),
        file_name=xlsx.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # PDF report
    pdf_report = results["report_pdf"]
    st.download_button(
        "⬇️ Download PDF report",
        data=pdf_report.read_bytes(),
        file_name=pdf_report.name,
        mime="application/pdf",
    )

    st.divider()
    st.caption("Raw CSV outputs (optional)")

    for fn in [
        "trip_level_edw_flags.csv",
        "trip_length_summary.csv",
        "edw_vs_day_summary.csv",
        "edw_by_length.csv",
        "edw_weighting_summary.csv",
        "edw_trip_ids.csv",
    ]:
        fp = out_dir / fn
        if fp.exists():
            st.download_button(
                f"Download {fn}",
                data=fp.read_bytes(),
                file_name=fp.name,
                mime="text/csv",
            )

st.caption(
    "Notes: EDW = any duty day touches 02:30–05:00 local (inclusive). "
    "Local hour comes from the number in parentheses ( ), minutes from the following Z time."
)
