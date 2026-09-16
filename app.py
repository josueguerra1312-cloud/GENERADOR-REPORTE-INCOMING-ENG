import streamlit as st
from pathlib import Path
from datetime import date
from report_generator import generate_report

st.set_page_config(page_title="Incoming Engine Report", layout="wide")
st.title("INCOMING ENGINE REPORT GENERATOR")
st.caption("Carga fotografías por sección y genera automáticamente un PowerPoint.")

sections = [
    "ENG DATA PLATES",
    "FWD AND AFTER",
    "FWD AND AFTER (CONE)",
    "FWD AND AFTER (FAN BLADES)",
    "FWD AND AFTER (RUBSTRIP PANEL, ACOUSTIC SEGMENT)",
    "FAN SECTION",
    "CORE",
    "UPPER AND LOWER",
    "UPPER AND LOWER (FWD MOUNT)",
    "EEC DATA PLATE & SOFTWARE PLATE",
    "DSU",
    "PHMU",
    "OIL TANK",
    "PLASTIC TARP INSIDE-LOWER AREA AND DESICCANTS",
    "Engine Stand & Cradle Data Plates",
    "SAFETY PINS ENG STAND",
    "HUMIDITY INDICATOR",
    "ENG COVERED",
    "TM VOI-TIC-111",
    "DISCREPANCIES",
]

with st.sidebar:
    st.header("Engine Information")
    pn = st.text_input("Part Number", "PW1133GA-JM")
    sn = st.text_input("Serial Number", "P800341")
    qc = st.text_input("QC", "GDL")
    report_date = st.date_input("Date", date.today())

uploads = {}
for section in sections:
    st.subheader(section)
    uploads[section] = st.file_uploader(
        f"Fotografías — {section}",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        key=section,
    )

st.divider()
if st.button("GENERATE POWERPOINT", type="primary", use_container_width=True):
    if not sn.strip():
        st.error("El Serial Number es obligatorio.")
        st.stop()

    with st.spinner("Generando reporte..."):
        output = generate_report(
            pn=pn.strip(),
            sn=sn.strip(),
            qc=qc.strip(),
            report_date=report_date.strftime("%d-%b-%Y").upper(),
            uploads=uploads,
        )

    st.success("Reporte generado correctamente.")
    st.download_button(
        "⬇️ Descargar PowerPoint",
        data=output.read_bytes(),
        file_name=output.name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )
