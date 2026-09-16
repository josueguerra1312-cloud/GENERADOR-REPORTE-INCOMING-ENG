import streamlit as st
from datetime import date
from report_generator import generate_report, SECTION_CONFIG

st.set_page_config(
    page_title="Incoming Engine Report | Volaris Style",
    page_icon="✈️",
    layout="wide",
)

st.markdown("""
<style>
    .main-title {font-size: 2rem; font-weight: 800; margin-bottom: 0.1rem;}
    .subtitle {color:#666; margin-bottom:1.2rem;}
    .section-title {font-size:1.05rem; font-weight:700; color:#8A005D;}
    div[data-testid="stFileUploader"] {border-radius: 8px;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">INCOMING ENGINE REPORT</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Automatic PowerPoint generator · Volaris-inspired technical layout</div>',
    unsafe_allow_html=True
)

with st.sidebar:
    st.header("ENGINE INFORMATION")
    pn = st.text_input("Part Number (PN)", "PW1133GA-JM")
    sn = st.text_input("Serial Number (SN)", "P800341")
    qc = st.text_input("QC", "GDL")
    report_date = st.date_input("Date", date.today())
    st.divider()
    st.caption("V2 · Dynamic photo layouts")

st.subheader("Photographs by section")
st.info(
    "Carga las fotografías de cada apartado. El generador conserva la proporción "
    "de las imágenes y crea diapositivas adicionales automáticamente."
)

uploads = {}
for key, config in SECTION_CONFIG.items():
    label = config["title"]
    with st.expander(label, expanded=False):
        uploads[key] = st.file_uploader(
            f"Fotografías — {label}",
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
            refs = st.text_input(
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
                "reference": refs,
                "photos": photos or [],
            })

st.divider()

if st.button("GENERATE POWERPOINT", type="primary", use_container_width=True):
    if not sn.strip():
        st.error("Serial Number (SN) is required.")
        st.stop()

    with st.spinner("Building Volaris-inspired technical report..."):
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
