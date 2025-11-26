# ===============================================================
#  TARTE BEAUTY  ·  LUXURY DASHBOARD
#  AI-Powered SKU Intake System
# ===============================================================

import streamlit as st
import sqlite3
import pytesseract
from PIL import Image
from docx import Document
import pandas as pd
import openpyxl
import os
import shutil
import platform
import subprocess
import base64

# ===============================================================
# GLOBAL CONFIG
# ===============================================================
EXCEL_PATH = "data/specs/sample_specs.xlsx"
EXCEL_SHEET_NAME = "Master Sheet - 12th Floor"
EXCEL_INSERT_ROW = 757

st.set_page_config(
    page_title="Tarte AI SKU System",
    page_icon="💜",
    layout="wide"
)

# ===============================================================
# LUXURY TARTE UI THEME
# ===============================================================

tarte_css = """
<style>

    /* GLOBAL BACKGROUND - soft gradient */
    .stApp {
    background: #c098ea !important;
    background-attachment: fixed;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }


    /* MAIN CONTENT WIDTH */
    div.block-container {
        max-width: 1180px;
        padding-top: 2.5rem;
        padding-bottom: 2.5rem;
        margin: 0 auto;
    }

    /* TOP TOOLBAR */
    header[data-testid="stHeader"] {
        background-color: rgba(0, 0, 0, 0);
    }

    /* SIDEBAR PANEL */
    section[data-testid="stSidebar"] {
    background: #e1daf8;
    border-right: 2px solid #8b55c9;
    padding-top: 2rem;
    }

    section[data-testid="stSidebar"] * {
        color: #240a3f !important;
        font-weight: 500;
    }

    /* Sidebar title */
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3 {
        color: #240a3f !important;
        font-weight: 800 !important;
    }

    /* Sidebar radio as pill nav */
    div[role="radiogroup"] > label {
        border-radius: 999px;
        padding: 0.3rem 0.9rem;
        margin-bottom: 0.25rem;
        transition: all 0.18s ease-in-out;
        border: 1px solid transparent;
    }

    div[role="radiogroup"] > label:hover {
        background-color: #f6edff40;
        border-color: #f6edff80;
    }

    /* Highlight selected option */
    div[role="radiogroup"] input[checked] + div {
        color: #240a3f !important;
        font-weight: 700;
    }

    /* OPEN EXCEL BUTTON IN SIDEBAR */
    section[data-testid="stSidebar"] div.stButton > button {
        width: 100%;
        background: linear-gradient(135deg, #f8d76a, #f0b94c);
        color: #3a2308 !important;
        border-radius: 999px !important;
        border: none !important;
        font-weight: 700 !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.18);
    }
    section[data-testid="stSidebar"] div.stButton > button:hover {
        background: linear-gradient(135deg, #ffe27c, #f6c764);
        transform: translateY(-1px);
    }

    /* LUX HEADER BANNER */
    .tarte-header {
        background: linear-gradient(90deg, #ffffffee, #f7efffff);
        padding: 26px 40px;
        border-radius: 26px;
        border: 1px solid #ebddff;
        box-shadow: 0 12px 30px rgba(115, 71, 153, 0.20);
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1.5rem;
        margin-bottom: 24px;
    }

    .tarte-header-left {
        text-align: left;
    }
    .tarte-header-title {
        font-size: 30px;
        font-weight: 800;
        color: #34114f;
        margin: 0;
    }
    .tarte-header-sub {
        margin: 4px 0 0 0;
        font-size: 14px;
        color: #62447f;
    }

    .tarte-header-pill {
        background: #f5ecff;
        border-radius: 999px;
        padding: 8px 16px;
        font-size: 13px;
        color: #6b3ea5;
        border: 1px solid #ecd9ff;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }

    /* WHITE CONTENT CARDS */
    .tarte-card {
        background: #ffffff;
        padding: 26px 24px;
        border-radius: 22px;
        border: 1px solid #ebddff;
        box-shadow: 0 8px 20px rgba(92, 54, 132, 0.10);
        margin-bottom: 26px;
    }

    /* SECTION TITLES */
    h1, h2, h3, h4, h5 {
        color: #34114f !important;
        font-weight: 750 !important;
    }
    p, li, label, span, div {
        color: #4b3569;
    }

    /* GOLD CTA BUTTONS (main area) */
    div.stButton > button {
        background: linear-gradient(135deg, #f3cc5d, #e4a93a) !important;
        color: #3a2308 !important;
        border-radius: 999px !important;
        padding: 0.5rem 1.3rem !important;
        border: none !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.18);
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #ffe27b, #f5c65c) !important;
        transform: translateY(-1px);
    }

    /* CHAT BUBBLES */
    .stChatMessage {
        background: #ffffffdd !important;
        border-radius: 12px !important;
        padding: 12px !important;
    }

</style>
"""
st.markdown(tarte_css, unsafe_allow_html=True)

# ===============================================================
# Helper: Open Excel
# ===============================================================
def open_excel_file(excel_path=EXCEL_PATH):
    try:
        system = platform.system()
        if system == "Darwin":
            subprocess.call(["open", excel_path])
        elif system == "Windows":
            os.startfile(excel_path)
        else:
            subprocess.call(["xdg-open", excel_path])
        st.success("📂 Excel file opened!")
    except Exception as e:
        st.error(f"❌ Could not open Excel file.\n\n{e}")

# ===============================================================
# Helper: image → base64 (for logo)
# ===============================================================
def load_image_base64(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

# ===============================================================
# 1. OCR AND DOCX EXTRACTION
# ===============================================================
def extract_text_from_image(image_file):
    img = Image.open(image_file)
    return pytesseract.image_to_string(img)

def extract_text_from_docx(docx_file):
    doc = Document(docx_file)
    result = []

    for para in doc.paragraphs:
        if para.text.strip():
            result.append(para.text.strip())

    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if len(cells) == 2:
                result.append(f"{cells[0]} {cells[1]}")
            elif cells:
                result.append(" | ".join(cells))

    return "\n".join(result)

# ===============================================================
# 2. PARSE FIELDS
# ===============================================================
def parse_vendor_doc(text):
    fields = {
        "Product Description": None,
        "Batch/Lot No.": None,
        "Date": None,
        "SKU": None,
        "Qty": None,
    }
    for line in text.splitlines():
        for key in fields.keys():
            if line.lower().startswith(key.lower()):
                parts = line.split(":", 1)
                if len(parts) == 2:
                    fields[key] = parts[1].strip()
    return fields

# ===============================================================
# 3. DATABASE
# ===============================================================
def init_db():
    conn = sqlite3.connect("sku_catalog.db")
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS sku_catalog (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_desc TEXT,
        batch_lot TEXT,
        date TEXT,
        sku TEXT,
        qty TEXT
    )
    """)
    conn.close()

def save_to_db(fields):
    conn = sqlite3.connect("sku_catalog.db")
    cur = conn.cursor()

    cur.execute("""
    SELECT COUNT(*) FROM sku_catalog WHERE sku = ? AND batch_lot = ?
    """, (fields["SKU"], fields["Batch/Lot No."]))
    exists = cur.fetchone()[0]

    if exists:
        st.warning(f"⚠️ Duplicate SKU {fields['SKU']} / Lot {fields['Batch/Lot No.']}")
    else:
        cur.execute("""
        INSERT INTO sku_catalog (product_desc, batch_lot, date, sku, qty)
        VALUES (?, ?, ?, ?, ?)
        """, (
            fields["Product Description"],
            fields["Batch/Lot No."],
            fields["Date"],
            fields["SKU"],
            fields["Qty"]
        ))
        conn.commit()
        st.success("✅ Saved to database!")

    conn.close()

def search_db(keyword):
    conn = sqlite3.connect("sku_catalog.db")
    df = pd.read_sql_query(f"""
        SELECT * FROM sku_catalog
        WHERE sku LIKE '%{keyword}%'
        OR batch_lot LIKE '%{keyword}%'
        OR product_desc LIKE '%{keyword}%'
    """, conn)
    conn.close()
    return df

def clear_database():
    conn = sqlite3.connect("sku_catalog.db")
    cur = conn.cursor()
    cur.execute("DELETE FROM sku_catalog")
    conn.commit()
    conn.close()

# ===============================================================
# 4. SAVE TO EXCEL
# ===============================================================
def save_to_excel(fields):
    if not os.path.exists(EXCEL_PATH):
        st.error(f"❌ Excel not found at {EXCEL_PATH}")
        return

    temp_path = EXCEL_PATH.replace(".xlsx", "_temp.xlsx")
    shutil.copyfile(EXCEL_PATH, temp_path)

    wb = openpyxl.load_workbook(temp_path)
    if EXCEL_SHEET_NAME not in wb.sheetnames:
        st.error("❌ Sheet name incorrect.")
        return

    ws = wb[EXCEL_SHEET_NAME]

    excel_headers = {}
    for col in range(1, ws.max_column + 1):
        header_value = ws.cell(row=1, column=col).value
        if header_value:
            excel_headers[header_value.lower()] = col

    mapping = {
        "product description": fields["Product Description"],
        "lot #": fields["Batch/Lot No."],
        "date": fields["Date"],
        "sku": fields["SKU"],
        "qty": fields["Qty"],
    }

    target_row = EXCEL_INSERT_ROW
    while ws.max_row < target_row:
        ws.append([])

    for key, value in mapping.items():
        if key in excel_headers:
            col = excel_headers[key]
            ws.cell(row=target_row, column=col, value=value)

    wb.save(temp_path)
    shutil.move(temp_path, EXCEL_PATH)
    st.success("💾 Saved to Excel successfully!")

# ===============================================================
# SIDEBAR NAVIGATION + LOGO
# ===============================================================
# Sidebar logo (Option 1 – logo above title)
logo_base64 = None
try:
    logo_base64 = load_image_base64("data/images/tarte_logo.png")
except FileNotFoundError:
    logo_base64 = None

if logo_base64:
    st.sidebar.markdown(
        f"""
        <div style="text-align:center; margin-bottom:10px;">
            <img src="data:image/png;base64,{logo_base64}"
                 style="width:140px; margin-top:-10px;">
        </div>
        """,
        unsafe_allow_html=True,
    )

st.sidebar.title(" Tarte SKU System")

page = st.sidebar.radio(
    "Navigate",
    ["Home", "Upload", "Camera", "Chatbot", "Database", "Excel Tools"]
)

if st.sidebar.button("📂 Open Excel File"):
    open_excel_file()

# ===============================================================
# HEADER BANNER (SHARED)
# ===============================================================
st.markdown(
    """
    <div class="tarte-header">
        <div class="tarte-header-left">
            <p class="tarte-header-title">Tarte AI-Powered SKU Intake</p>
            <p class="tarte-header-sub">
                Designed for production retains, lab samples, and go to market traceability.
            </p>
        </div>
        <div>
            <span class="tarte-header-pill">
                🧾 Auto intake · ✅ Duplicate safe · 📊 Excel synced
            </span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ===============================================================
# HOME PAGE
# ===============================================================
if page == "Home":
    st.subheader("✨ Welcome")

    st.write(
        """
This dashboard automates the workflow for cataloging Tarte production retains:

• Upload DOCX or image files  
• Capture new retains with the camera  
• Extract and parse core SKU fields  
• Store records in a duplicate safe SQLite database  
• Sync rows directly into the master Excel tracker  
• Use the chatbot to trigger actions or search data  
        """
    )

# ===============================================================
# UPLOAD PAGE
# ===============================================================
elif page == "Upload":

    st.markdown('<div class="tarte-card">', unsafe_allow_html=True)
    st.header("📄 Upload Vendor Documents")

    files = st.file_uploader(
        "Upload one or more files",
        type=["docx", "jpg", "jpeg", "png"],
        accept_multiple_files=True,
        help="Upload retain specs, artwork proofs, or vendor docs."
    )

    collected_fields = []

    if files:
        st.success(f"{len(files)} files uploaded.")

        for idx, f in enumerate(files, start=1):
            st.markdown("---")
            st.subheader(f"📄 File {idx}: {f.name}")

            text = (
                extract_text_from_docx(f)
                if f.name.lower().endswith(".docx")
                else extract_text_from_image(f)
            )

            with st.expander("🔎 Extracted text preview"):
                st.text(text)

            fields = parse_vendor_doc(text)
            st.json(fields)
            collected_fields.append(fields)

            if st.button(f"💾 Save {f.name}", key=f"save_{idx}"):
                save_to_db(fields)
                save_to_excel(fields)

        if collected_fields and st.button("🔥 Process all files"):
            for flds in collected_fields:
                save_to_db(flds)
                save_to_excel(flds)
            st.success("All files processed and synced.")

    st.markdown("</div>", unsafe_allow_html=True)

# ===============================================================
# CAMERA PAGE
# ===============================================================
elif page == "Camera":

    st.markdown('<div class="tarte-card">', unsafe_allow_html=True)
    st.header("📸 Capture Vendor Document")

    photo = st.camera_input("Take a photo of the retain or vendor sheet")

    if photo:
        text = extract_text_from_image(photo)
        with st.expander("🔎 Extracted text preview"):
            st.text(text)

        fields = parse_vendor_doc(text)
        st.json(fields)

        if st.button("💾 Save from camera"):
            save_to_db(fields)
            save_to_excel(fields)

    st.markdown("</div>", unsafe_allow_html=True)

# ===============================================================
# CHATBOT PAGE
# ===============================================================
elif page == "Chatbot":

    st.markdown('<div class="tarte-card">', unsafe_allow_html=True)
    st.header("🤖 Chatbot Mode")

    if "messages" not in st.session_state:
        st.session_state.messages = []

    for m in st.session_state.messages:
        with st.chat_message(m["role"]):
            st.write(m["content"])

    user_input = st.chat_input("Ask me to process latest, search, or show database")

    if user_input:

        st.session_state.messages.append({"role": "user", "content": user_input})

        if "process latest" in user_input.lower():
            folder = "data/specs/"
            files = [
                os.path.join(folder, f)
                for f in os.listdir(folder)
                if f.lower().endswith((".docx", ".jpg", ".jpeg", ".png"))
            ]

            if not files:
                reply = "❌ No files found in data/specs."
            else:
                latest = max(files, key=os.path.getmtime)
                text = (
                    extract_text_from_docx(latest)
                    if latest.lower().endswith(".docx")
                    else extract_text_from_image(latest)
                )
                fields = parse_vendor_doc(text)
                save_to_db(fields)
                save_to_excel(fields)
                reply = f"Processed and saved: {os.path.basename(latest)}"

        elif "search" in user_input.lower():
            term = user_input.split("search", 1)[1].strip()
            df = search_db(term)
            st.dataframe(df)
            reply = f"Search results for '{term}'."

        elif "show database" in user_input.lower():
            conn = sqlite3.connect("sku_catalog.db")
            df = pd.read_sql_query("SELECT * FROM sku_catalog", conn)
            conn.close()
            st.dataframe(df)
            reply = "Full database view."

        elif "clear database" in user_input.lower():
            clear_database()
            reply = "Database cleared."

        else:
            reply = "Unknown command. Try: process latest, search <term>, show database, or clear database."

        st.session_state.messages.append({"role": "assistant", "content": reply})

    st.markdown("</div>", unsafe_allow_html=True)

# ===============================================================
# DATABASE PAGE
# ===============================================================
elif page == "Database":

    st.markdown('<div class="tarte-card">', unsafe_allow_html=True)
    st.header("📊 Database Viewer")

    init_db()

    search = st.text_input("Search SKU, lot, or product description")

    if search:
        df = search_db(search)
        st.dataframe(df)
    else:
        conn = sqlite3.connect("sku_catalog.db")
        df = pd.read_sql_query("SELECT * FROM sku_catalog", conn)
        conn.close()
        st.dataframe(df)

    if st.button("🧹 Clear database"):
        clear_database()
        st.success("Database cleared.")

    st.markdown("</div>", unsafe_allow_html=True)

# ===============================================================
# EXCEL TOOLS
# ===============================================================
elif page == "Excel Tools":

    st.markdown('<div class="tarte-card">', unsafe_allow_html=True)
    st.header("📂 Excel Tools")

    st.write(f"Excel file: `{EXCEL_PATH}`")
    st.write(f"Sheet name: `{EXCEL_SHEET_NAME}`")

    if st.button("📂 Open Excel file"):
        open_excel_file()

    st.markdown("---")
    st.write("Tip: close the workbook before running updates from this dashboard.")

    st.markdown("</div>", unsafe_allow_html=True)
