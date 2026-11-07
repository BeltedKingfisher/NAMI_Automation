"""Minimal upload UI.  Run with:  streamlit run app.py"""
import streamlit as st

from pipeline import process_uploads, rows_to_xlsx

st.title("Survey digitizer")
st.write("Upload scanned surveys (3-page PDFs, or images in page order) and get an Excel file back.")

uploads = st.file_uploader(
    "Survey files", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True
)

if uploads and st.button("Process"):
    bar = st.progress(0.0)
    files = [(u.name, u.getvalue()) for u in uploads]
    try:
        rows = process_uploads(files, progress=bar.progress)
    except NotImplementedError as e:
        st.error(f"OCR not wired up yet: {e}")
        st.stop()

    flagged = sum(1 for r in rows if r.flags)
    st.success(f"Processed {len(rows)} surveys. {flagged} need review (highlighted in yellow).")
    st.download_button(
        "Download Excel",
        data=rows_to_xlsx(rows),
        file_name="survey_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )