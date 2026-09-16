from pathlib import Path
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image

SECTIONS = [
    "ENG DATA PLATES", "FWD AND AFTER", "FWD AND AFTER (CONE)",
    "FWD AND AFTER (FAN BLADES)",
    "FWD AND AFTER (RUBSTRIP PANEL, ACOUSTIC SEGMENT)",
    "FAN SECTION", "CORE", "UPPER AND LOWER",
    "UPPER AND LOWER (FWD MOUNT)", "EEC DATA PLATE & SOFTWARE PLATE",
    "DSU", "PHMU", "OIL TANK",
    "PLASTIC TARP INSIDE-LOWER AREA AND DESICCANTS",
    "Engine Stand & Cradle Data Plates", "SAFETY PINS ENG STAND",
    "HUMIDITY INDICATOR", "ENG COVERED", "TM VOI-TIC-111", "DISCREPANCIES",
]

OUT = Path("output")
OUT.mkdir(exist_ok=True)

def _add_cover(prs, pn, sn, qc, report_date):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tx = slide.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(11.7), Inches(3.8))
    tf = tx.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = "ENG INCOMING"
    p.font.size = Pt(30)
    p.font.bold = True
    for label, value in [
        ("PN", pn), ("SN", sn), ("QC", qc), ("DATE", report_date)
    ]:
        p = tf.add_paragraph()
        p.text = f"{label}: {value}"
        p.font.size = Pt(20)

def _add_section_slide(prs, title, image_files, per_page=4):
    if not image_files:
        return

    for start in range(0, len(image_files), per_page):
        batch = image_files[start:start + per_page]
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        header = slide.shapes.add_textbox(Inches(0.35), Inches(0.15), Inches(12.6), Inches(0.55))
        p = header.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(17)
        p.font.bold = True

        positions = [
            (0.45, 0.85, 6.0, 3.0),
            (6.85, 0.85, 6.0, 3.0),
            (0.45, 4.05, 6.0, 3.0),
            (6.85, 4.05, 6.0, 3.0),
        ]

        for img, (x, y, w, h) in zip(batch, positions):
            data = img.getvalue() if hasattr(img, "getvalue") else Path(img).read_bytes()
            with Image.open(BytesIO(data)) as im:
                ratio = im.width / im.height
            box_ratio = w / h
            if ratio > box_ratio:
                width = w
                height = w / ratio
                x2, y2 = x, y + (h - height) / 2
            else:
                height = h
                width = h * ratio
                x2, y2 = x + (w - width) / 2, y
            slide.shapes.add_picture(BytesIO(data), Inches(x2), Inches(y2),
                                     width=Inches(width), height=Inches(height))

def generate_report(pn, sn, qc, report_date, uploads):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    _add_cover(prs, pn, sn, qc, report_date)

    for section in SECTIONS:
        files = uploads.get(section, [])
        _add_section_slide(prs, section, files)

    path = OUT / f"INCOMING_{sn}.pptx"
    prs.save(path)
    return path
