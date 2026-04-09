from __future__ import annotations

import json
import re
import uuid
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any

import streamlit as st
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage


APP_TITLE = "Inspection Photo Report Tool"
TEMPLATE_FILENAME = "WIR Photo Tool V1.0.xlsx"
TARGET_SHEET = "TOOL"
MAX_UPLOADS = 20
RETENTION_HOURS = 24

# Excel display size based on the user's screenshot
IMAGE_WIDTH_PX = int(round(2.49 * 96))   # ~239 px
IMAGE_HEIGHT_PX = int(round(1.87 * 96))  # ~180 px

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "app_data"
REPORT_DIR = DATA_DIR / "reports"
TMP_DIR = DATA_DIR / "tmp_images"
INDEX_FILE = DATA_DIR / "report_index.json"
TEMPLATE_PATH = BASE_DIR / TEMPLATE_FILENAME

IMAGE_CELLS = [
    "B12", "G12", "G21", "B21", "B31", "G31",
    "B64", "G64", "G73", "B73", "B83", "G83",
    "B116", "G116", "G125", "B125", "B135", "G135",
    "B168", "G168", "G177", "B177", "B187", "G187",
    "B220", "G220", "G229", "B229", "B239", "G239",
    "B272", "G272", "G281", "B281", "B291", "G291",
]

ALLOWED_SUFFIXES = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    TMP_DIR.mkdir(parents=True, exist_ok=True)
    if not INDEX_FILE.exists():
        INDEX_FILE.write_text("[]", encoding="utf-8")


def load_index() -> list[dict[str, Any]]:
    ensure_dirs()
    try:
        data = json.loads(INDEX_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []


def save_index(items: list[dict[str, Any]]) -> None:
    INDEX_FILE.write_text(json.dumps(items, indent=2, ensure_ascii=False), encoding="utf-8")


def parse_dt(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except Exception:
        return None


def now_local() -> datetime:
    return datetime.now()


def cleanup_expired() -> None:
    current = now_local()
    kept: list[dict[str, Any]] = []

    for item in load_index():
        report_path = REPORT_DIR / item.get("stored_name", "")
        expires_at = parse_dt(item.get("expires_at"))
        downloaded = bool(item.get("downloaded"))

        should_remove = False
        if not report_path.exists():
            should_remove = True
        elif expires_at and current >= expires_at:
            should_remove = True
        elif downloaded:
            should_remove = True

        if should_remove:
            try:
                report_path.unlink(missing_ok=True)
            except Exception:
                pass
        else:
            kept.append(item)

    save_index(kept)

    for temp_file in TMP_DIR.glob("*"):
        try:
            modified = datetime.fromtimestamp(temp_file.stat().st_mtime)
            if current - modified > timedelta(hours=RETENTION_HOURS):
                temp_file.unlink(missing_ok=True)
        except Exception:
            pass


def sanitize_filename_component(text: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]', "", text or "")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned or "Inspector"


def build_output_filename(inspector_name: str) -> str:
    ts = now_local()
    date_part = ts.strftime("%Y-%m-%d")
    time_part = ts.strftime("%I%M")  # 4:58pm -> 0458
    safe_name = sanitize_filename_component(inspector_name)
    return f"Inspection_Report_{date_part}_{time_part} ({safe_name}).xlsx"


def register_report(display_name: str, stored_name: str, inspector_name: str, equipment_room: str) -> None:
    current = now_local()
    index_data = load_index()
    index_data.insert(0, {
        "id": uuid.uuid4().hex,
        "display_name": display_name,
        "stored_name": stored_name,
        "inspector_name": inspector_name,
        "equipment_room": equipment_room,
        "created_at": current.isoformat(timespec="seconds"),
        "expires_at": (current + timedelta(hours=RETENTION_HOURS)).isoformat(timespec="seconds"),
        "downloaded": False,
    })
    save_index(index_data)


def mark_report_downloaded(report_id: str) -> None:
    index_data = load_index()
    kept: list[dict[str, Any]] = []

    for item in index_data:
        if item.get("id") == report_id:
            report_path = REPORT_DIR / item.get("stored_name", "")
            try:
                report_path.unlink(missing_ok=True)
            except Exception:
                pass
            continue
        kept.append(item)

    save_index(kept)


def get_pending_reports() -> list[dict[str, Any]]:
    cleanup_expired()
    pending: list[dict[str, Any]] = []
    for item in load_index():
        report_path = REPORT_DIR / item.get("stored_name", "")
        if not item.get("downloaded") and report_path.exists():
            pending.append(item)
    return pending


def validate_uploads(uploaded_files: list[Any]) -> None:
    if not uploaded_files:
        raise ValueError("Please upload at least 1 image.")
    if len(uploaded_files) > MAX_UPLOADS:
        raise ValueError(f"You can upload maximum {MAX_UPLOADS} images.")
    for file in uploaded_files:
        suffix = Path(file.name).suffix.lower()
        if suffix not in ALLOWED_SUFFIXES:
            raise ValueError(f"Unsupported file type: {file.name}")


def save_uploaded_file_temporarily(uploaded_file: Any, seq_no: int) -> Path:
    suffix = Path(uploaded_file.name).suffix.lower()
    temp_name = f"{uuid.uuid4().hex}_{seq_no:02d}{suffix}"
    temp_path = TMP_DIR / temp_name
    temp_path.write_bytes(uploaded_file.getbuffer())
    return temp_path


def generate_report_bytes(inspector_name: str, equipment_room: str, uploaded_files: list[Any]) -> bytes:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template file not found: {TEMPLATE_FILENAME}")

    validate_uploads(uploaded_files)

    wb = load_workbook(TEMPLATE_PATH)
    if TARGET_SHEET not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Sheet '{TARGET_SHEET}' not found in template.")

    ws = wb[TARGET_SHEET]
    ws["G8"] = equipment_room.strip()

    temp_paths: list[Path] = []

    try:
        for idx, uploaded in enumerate(uploaded_files[:len(IMAGE_CELLS)], start=1):
            temp_path = save_uploaded_file_temporarily(uploaded, idx)
            temp_paths.append(temp_path)

            img = XLImage(str(temp_path))
            # Keep original image data; only set display size in Excel.
            img.width = IMAGE_WIDTH_PX
            img.height = IMAGE_HEIGHT_PX
            ws.add_image(img, IMAGE_CELLS[idx - 1])

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()
    finally:
        try:
            wb.close()
        except Exception:
            pass
        for temp_path in temp_paths:
            try:
                temp_path.unlink(missing_ok=True)
            except Exception:
                pass


def save_report_to_disk(file_name: str, report_bytes: bytes) -> str:
    stored_name = f"{uuid.uuid4().hex}_{file_name}"
    report_path = REPORT_DIR / stored_name
    report_path.write_bytes(report_bytes)
    return stored_name


def format_dt(value: str | None) -> str:
    dt = parse_dt(value)
    if not dt:
        return "-"
    return dt.strftime("%Y-%m-%d %H:%M")


def format_size(num_bytes: int) -> str:
    kb = num_bytes / 1024
    if kb < 1024:
        return f"{kb:.1f} KB"
    return f"{kb / 1024:.2f} MB"


st.set_page_config(page_title=APP_TITLE, page_icon="📷", layout="wide")
ensure_dirs()
cleanup_expired()

if "success_message" not in st.session_state:
    st.session_state["success_message"] = ""

st.markdown(
    """
    <style>
    .main > div {padding-top: 2rem;}
    .block-container {max-width: 1200px;}
    div[data-testid="stForm"] {
        border: 1px solid rgba(255,255,255,0.12);
        border-radius: 18px;
        padding: 1.25rem 1.25rem 1rem 1.25rem;
        background: rgba(255,255,255,0.02);
    }
    .status-box {
        padding: 1rem 1.2rem;
        border-radius: 14px;
        margin-bottom: 1rem;
        border: 1px solid rgba(255,255,255,0.08);
        background: linear-gradient(90deg, rgba(26, 128, 74, 0.85), rgba(24, 90, 48, 0.85));
        color: white;
        font-size: 1.05rem;
    }
    .pending-card {
        padding: 1rem 1.1rem;
        border-radius: 14px;
        border: 1px solid rgba(255,255,255,0.08);
        background: rgba(255,255,255,0.02);
        margin-bottom: 0.8rem;
    }
    .small-muted {
        opacity: 0.8;
        font-size: 0.92rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📷 Inspection Photo Report Tool")
st.caption("Upload inspection photos and generate an Excel report from your template.")

if TEMPLATE_PATH.exists():
    st.markdown(f'<div class="status-box">Template loaded: <strong>{TEMPLATE_FILENAME}</strong></div>', unsafe_allow_html=True)
else:
    st.error(f"Template file not found: {TEMPLATE_FILENAME}")

if st.session_state["success_message"]:
    st.success(st.session_state["success_message"])
    st.session_state["success_message"] = ""

st.subheader("Generate new report")

with st.form("report_form", clear_on_submit=True):
    inspector_name = st.text_input("Inspector Name", value="", placeholder="Enter inspector name")
    equipment_room = st.text_input("Equipment / Room", value="", placeholder="Enter equipment or room")
    uploaded_files = st.file_uploader(
        f"Select images (maximum {MAX_UPLOADS})",
        type=["jpg", "jpeg", "png", "bmp", "webp"],
        accept_multiple_files=True,
        help="Upload up to 20 images.",
    )

    selected_count = len(uploaded_files) if uploaded_files else 0
    st.caption(f"Selected images: {selected_count}/{MAX_UPLOADS}")

    submitted = st.form_submit_button("Generate Report", use_container_width=False, type="primary")

if submitted:
    try:
        if not inspector_name.strip():
            raise ValueError("Please enter Inspector Name.")
        if not equipment_room.strip():
            raise ValueError("Please enter Equipment / Room.")
        if not uploaded_files:
            raise ValueError("Please upload at least 1 image.")
        if len(uploaded_files) > MAX_UPLOADS:
            raise ValueError(f"You can upload maximum {MAX_UPLOADS} images.")

        output_name = build_output_filename(inspector_name)
        report_bytes = generate_report_bytes(inspector_name, equipment_room, uploaded_files)
        stored_name = save_report_to_disk(output_name, report_bytes)
        register_report(output_name, stored_name, inspector_name.strip(), equipment_room.strip())

        st.session_state["success_message"] = f"Report generated successfully: {output_name}"
        st.rerun()

    except Exception as exc:
        st.error(f"Failed to generate report: {exc}")

st.divider()
st.subheader("Pending reports not downloaded yet")

pending_reports = get_pending_reports()

if not pending_reports:
    st.info("No pending files. Downloaded files are removed immediately. Undownloaded files stay for up to 24 hours.")
else:
    st.caption("These files will be automatically deleted after 24 hours if they are not downloaded.")

    for item in pending_reports:
        report_path = REPORT_DIR / item["stored_name"]
        if not report_path.exists():
            continue

        col1, col2 = st.columns([4, 1])
        with col1:
            st.markdown(
                f"""
                <div class="pending-card">
                    <div><strong>{item['display_name']}</strong></div>
                    <div class="small-muted">
                        Inspector: {item.get('inspector_name', '-')}<br>
                        Equipment / Room: {item.get('equipment_room', '-')}<br>
                        Created: {format_dt(item.get('created_at'))}<br>
                        Expires: {format_dt(item.get('expires_at'))}<br>
                        Size: {format_size(report_path.stat().st_size)}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with col2:
            file_bytes = report_path.read_bytes()
            clicked = st.download_button(
                label="Download",
                data=file_bytes,
                file_name=item["display_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{item['id']}",
                use_container_width=True,
            )
            if clicked:
                mark_report_downloaded(item["id"])
                st.rerun()
