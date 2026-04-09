import io
import os
import re
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageOps

# =========================
# Config
# =========================
APP_TITLE = "Inspection Photo Report Tool"
TEMPLATE_FILENAME = "WIR Photo Tool V1.0.xlsx"
TARGET_SHEET = "TOOL"
MAX_IMAGES = 20

# Excel picture size from user's requirement
IMG_WIDTH_IN = 2.49
IMG_HEIGHT_IN = 1.87
EXCEL_DPI = 96
IMG_WIDTH_PX = int(round(IMG_WIDTH_IN * EXCEL_DPI))   # ~239 px
IMG_HEIGHT_PX = int(round(IMG_HEIGHT_IN * EXCEL_DPI)) # ~180 px

# 6 image slots per page, 6 pages total in template
IMAGE_CELLS = [
    # Page 1
    "B12", "G12", "G21", "B21", "B31", "G31",
    # Page 2
    "B64", "G64", "G73", "B73", "B83", "G83",
    # Page 3
    "B116", "G116", "G125", "B125", "B135", "G135",
    # Page 4
    "B168", "G168", "G177", "B177", "B187", "G187",
    # Page 5
    "B220", "G220", "G229", "B229", "B239", "G239",
    # Page 6
    "B272", "G272", "G281", "B281", "B291", "G291",
]


# =========================
# Helpers
# =========================
def find_template_path() -> Path:
    """Find the Excel template in common locations."""
    candidates = [
        Path(__file__).with_name(TEMPLATE_FILENAME),
        Path.cwd() / TEMPLATE_FILENAME,
        Path("/mnt/data") / TEMPLATE_FILENAME,
    ]
    for path in candidates:
        if path.exists():
            return path
    raise FileNotFoundError(
        f"Template file '{TEMPLATE_FILENAME}' not found. Put it in the same folder as app.py."
    )


def sanitize_filename_part(name: str) -> str:
    """Keep filename safe for Windows/macOS/Linux."""
    cleaned = (name or "").strip()
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned[:80] if cleaned else "Inspector"



def build_output_filename(inspector_name: str) -> str:
    date_str = datetime.now().strftime("%Y-%m-%d")
    safe_name = sanitize_filename_part(inspector_name)
    return f"Inspection_Report_{date_str} ({safe_name}).xlsx"



def resize_with_padding(uploaded_file, width_px: int, height_px: int) -> Image.Image:
    """Resize image to fit exact target box while preserving aspect ratio.
    Pads with white background so all images appear same size in Excel.
    """
    img = Image.open(io.BytesIO(uploaded_file.getvalue()))
    img = ImageOps.exif_transpose(img)

    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGB")

    # Keep transparency clean if source has alpha
    if img.mode == "RGBA":
        background = Image.new("RGB", img.size, "white")
        background.paste(img, mask=img.getchannel("A"))
        img = background

    # Fit inside target and pad to exact size
    fitted = ImageOps.contain(img, (width_px, height_px), Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", (width_px, height_px), "white")
    x = (width_px - fitted.width) // 2
    y = (height_px - fitted.height) // 2
    canvas.paste(fitted, (x, y))
    return canvas



def generate_report_bytes(inspector_name: str, equipment_room: str, uploaded_images) -> bytes:
    template_path = find_template_path()
    wb = load_workbook(template_path)

    if TARGET_SHEET not in wb.sheetnames:
        raise ValueError(f"Sheet '{TARGET_SHEET}' not found in template.")

    ws = wb[TARGET_SHEET]

    # Copy Equipment / Room name to TOOL!G8
    ws["G8"] = equipment_room.strip()

    # Insert images in order
    for idx, uploaded in enumerate(uploaded_images):
        if idx >= len(IMAGE_CELLS):
            break

        cell = IMAGE_CELLS[idx]
        prepared = resize_with_padding(uploaded, IMG_WIDTH_PX, IMG_HEIGHT_PX)
        xl_img = XLImage(prepared)
        xl_img.width = IMG_WIDTH_PX
        xl_img.height = IMG_HEIGHT_PX
        ws.add_image(xl_img, cell)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# =========================
# UI
# =========================
st.set_page_config(page_title=APP_TITLE, page_icon="📸", layout="centered")
st.title("📸 Inspection Photo Report Tool")
st.caption("Upload inspection photos and generate an Excel report from your template.")

# Template check
try:
    template_path = find_template_path()
    st.success(f"Template loaded: {template_path.name}")
except Exception as exc:
    st.error(str(exc))
    st.stop()

with st.form("inspection_form"):
    inspector_name = st.text_input("Inspector Name", placeholder="Enter inspector name")
    equipment_room = st.text_input("Equipment / Room", placeholder="Enter equipment or room name")
    uploaded_images = st.file_uploader(
        "Select images (maximum 20)",
        type=["jpg", "jpeg", "png", "bmp", "webp"],
        accept_multiple_files=True,
        help="Images will be inserted into sheet TOOL in the order you upload them.",
    )
    submitted = st.form_submit_button("Generate Report")

if submitted:
    errors = []

    if not inspector_name.strip():
        errors.append("Please enter Inspector Name.")

    if not equipment_room.strip():
        errors.append("Please enter Equipment / Room.")

    if not uploaded_images:
        errors.append("Please upload at least 1 image.")

    if uploaded_images and len(uploaded_images) > MAX_IMAGES:
        errors.append(f"You uploaded {len(uploaded_images)} images. Maximum allowed is {MAX_IMAGES}.")

    if errors:
        for err in errors:
            st.error(err)
    else:
        with st.spinner("Generating Excel report..."):
            try:
                report_bytes = generate_report_bytes(
                    inspector_name=inspector_name,
                    equipment_room=equipment_room,
                    uploaded_images=uploaded_images,
                )
                filename = build_output_filename(inspector_name)

                st.success("Report generated successfully.")
                st.download_button(
                    label="⬇️ Download Inspection Report",
                    data=report_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.info(
                    "This version does not store uploaded images or generated Excel files in a database. "
                    "The report is created in memory only, so there is normally nothing to keep or clean for 24 hours."
                )

                with st.expander("Preview of image placement logic"):
                    for i, cell in enumerate(IMAGE_CELLS[:MAX_IMAGES], start=1):
                        st.write(f"Image {i} → {TARGET_SHEET}!{cell}")

            except Exception as exc:
                st.exception(exc)

st.markdown("---")
st.markdown(
    "**Current logic:** Equipment / Room → `TOOL!G8`, images inserted in sequence, file name format → "
    "`Inspection_Report_YYYY-MM-DD (InspectorName).xlsx`"
)
