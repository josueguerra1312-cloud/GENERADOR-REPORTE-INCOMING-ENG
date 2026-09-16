import streamlit as st
from datetime import date
from pathlib import Path
import importlib.util

# Load report_generator explicitly from the same directory as this file.
# This avoids Streamlit Cloud working-directory/import-path issues.
BASE_DIR = Path(__file__).resolve().parent
GENERATOR_PATH = BASE_DIR / "report_generator.py"

if not GENERATOR_PATH.exists():
    st.error(
        "No se encontró report_generator.py en la raíz del repositorio. "
        "Sube app.py y report_generator.py al mismo nivel en GitHub."
    )
    st.stop()

spec = importlib.util.spec_from_file_location("incoming_report_generator", GENERATOR_PATH)
if spec is None or spec.loader is None:
    st.error("No fue posible cargar report_generator.py.")
    st.stop()

report_generator = importlib.util.module_from_spec(spec)
try:
    spec.loader.exec_module(report_generator)
except Exception as exc:
    st.error("No fue posible cargar el generador de reportes.")
    st.code(f"{type(exc).__name__}: {exc}")
    st.stop()

generate_report = report_generator.generate_report
SECTION_CONFIG = report_generator.SECTION_CONFIG

st.set_page_config(
    page_title="Incoming Engine Report | Volaris Style",
    page_icon="✈️",
    layout="wide",
)

st.markdown("""
<style>
.main-title {font-size: 2rem; font-weight: 800; margin-bottom: .1rem;}
.subtitle {color:#666; margin-bottom:1.2rem;}
[data-testid="stFileUploader"] {border-radius:8px;}
</style>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="main-title">INCOMING ENGINE REPORT</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="subtitle">Automatic PowerPoint generator · Volaris-inspired technical layout</div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("ENGINE INFORMATION")
    pn = st.text_input("Part Number (PN)", "PW1133GA-JM")
    sn = st.text_input("Serial Number (SN)", "P800341")
    qc = st.text_input("QC", "GDL")
    report_date = st.date_input("Date", date.today())
    st.divider()
    st.caption("V2.1 · Streamlit Cloud compatible")

st.subheader("Photographs by section")
st.info(
    "Carga las fotografías de cada apartado. Se conserva la proporción "
    "de cada imagen y se crean páginas adicionales automáticamente."
)

uploads = {}
for key, config in SECTION_CONFIG.items():
    with st.expander(config["title"], expanded=False):
        uploads[key] = st.file_uploader(
            f"Fotografías — {config['title']}",
            type=["jpg", "jpeg", "png", "webp"],
            accept_multiple_files=True,
            key=f"upload_{key}",
        )
        if uploads[key]:
            st.caption(f"{len(uploads[key])} fotografía(s) seleccionada(s).")

st.divider()
st.subheader("Discrepancies")

discrepancies = []
for i in range(1, 6):
    with st.expander(f"Discrepancy {i}", expanded=(i == 1)):
        enabled = st.checkbox("Include this discrepancy", key=f"disc_enabled_{i}")
        if enabled:
            description = st.text_input(
                "Description",
                key=f"disc_desc_{i}",
                placeholder="Ej. OIL TUBE (LR21) WITH DENT",
            )
            finding = st.text_area(
                "Finding / Observation",
                key=f"disc_finding_{i}",
                placeholder="Ej. DENT FREE OF SHARP AND CORNERS WITHIN LIMIT",
            )
            reference = st.text_input(
                "Reference / disposition",
                key=f"disc_ref_{i}",
                placeholder="Opcional",
            )
            photos = st.file_uploader(
                "Discrepancy photographs",
                type=["jpg", "jpeg", "png", "webp"],
                accept_multiple_files=True,
                key=f"disc_photos_{i}",
            )
            discrepancies.append({
                "description": description,
                "finding": finding,
                "reference": reference,
                "photos": photos or [],
            })

st.divider()

if st.button("GENERATE POWERPOINT", type="primary", use_container_width=True):
    if not sn.strip():
        st.error("Serial Number (SN) is required.")
        st.stop()

    try:
        with st.spinner("Building report..."):
            output = generate_report(
                pn=pn.strip(),
                sn=sn.strip(),
                qc=qc.strip(),
                report_date=report_date.strftime("%d-%b-%Y").upper(),
                uploads=uploads,
                discrepancies=discrepancies,
            )

        st.success(f"Report generated: {output.name}")
        st.download_button(
            "⬇️ DOWNLOAD POWERPOINT",
            data=output.read_bytes(),
            file_name=output.name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
    except Exception as exc:
        st.error("Ocurrió un error durante la generación del PowerPoint.")
        st.exception(exc)
