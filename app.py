from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st

from automate_imaging_opd import process_file

st.set_page_config(page_title="Imaging and OPD Automation", layout="centered")
st.title("Imaging and OPD Template Automation")
st.write("Upload a facility workbook. The completed file keeps the uploaded filename exactly.")

facility_upload = st.file_uploader("Facility Excel file", type=["xlsx"])
template_upload = st.file_uploader(
    "Template Excel file",
    type=["xlsx"],
    help="Upload Template Imaging and OPD.xlsx",
)

if facility_upload and template_upload:
    with TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        facility_path = tmp_dir / facility_upload.name
        template_path = tmp_dir / template_upload.name
        facility_path.write_bytes(facility_upload.getvalue())
        template_path.write_bytes(template_upload.getvalue())

        try:
            output_path = process_file(facility_path, template_path, tmp_dir / "output")
            output_bytes = output_path.read_bytes()
            st.success("The completed workbook is ready.")
            st.download_button(
                "Download completed workbook",
                data=output_bytes,
                file_name=facility_upload.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(str(exc))
