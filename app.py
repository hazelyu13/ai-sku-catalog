# ===============================================================
#  TARTE BEAUTY — FULL REDESIGN STREAMLIT APPLICATION
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
# TARTE PURPLE UI THEME — CSS INJECTION
# ===============================================================

tarte_css = """
<style>

    <style>
    /* FULL GRADIENT BACKGROUND */
    stApp { background: linear-gradient(180deg, #6d2ea6 0%, #b088d1 20%, #f5eafe 150%) !important; background-attachment: fixed; }

    /* SIDEBAR — matched gradient */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #6d2ea6 0%, #b088d1 20%, #f5eafe 150%) !important;
        background-attachment: fixed;
        padding-top: 2rem;
    }

    /* GOLD VERTICAL LINE ON SIDEBAR EDGE */
    section[data-testid="stSidebar"] > div:first-child {
        border-right: 3px solid #d4af37 !important; /* tarte gold */
        padding-right: 10px;
    }

    /* HEADER CARD */
    .tarte-header {
        background: #FFFFFF20;
        padding: 20px;
        border-radius: 18px;
        text-align: center;
        border: 1px solid #FFFFFF40;
        box-shadow: 0 4px 14px rgba(0,0,0,0.18);
        margin-bottom: 20px;
    }

    .tarte-header h2 {
        color: #ffffff !important;
        font-size: 32px;
        margin-bottom: 2px;
        font-weight: 700;
    }

    .tarte-header p {
        color: #f8eefe !important;
        font-size: 16px;
        margin-top: -5px;
    }

    /* CONTENT CARDS */
    .tarte-card {
        background: #ffffffdd;
        padding: 22px;
        border-radius: 18px;
        border: 1px solid #e8d8ff;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        margin-bottom: 22px;
    }

    /* GOLD BUTTONS */
    div.stButton > button {
        background-color: #cda349 !important;
        color: white !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 10px 20px !important;
        font-weight: bold !important;
        font-size: 15px !important;
    }

    div.stButton > button:hover {
        background-color: #e0b863 !important;
        color: #2B0E41 !important;
        transform: scale(1.02);
    }

    /* CHAT BUBBLES */
    .stChatMessage {
        background: #ffffffcc !important;
        backdrop-filter: blur(8px);
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
# 3. DATABASE (INIT + SAVE + SEARCH + CLEAR)
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
        st.warning(f"⚠️ Duplicate detected: SKU {fields['SKU']} / Lot {fields['Batch/Lot No.']}")
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
        st.error("❌ Excel file not found.")
        return

    temp_path = EXCEL_PATH.replace(".xlsx", "_temp.xlsx")
    shutil.copyfile(EXCEL_PATH, temp_path)

    wb = openpyxl.load_workbook(temp_path)
    if EXCEL_SHEET_NAME not in wb.sheetnames:
        st.error(f"❌ Sheet not found: {EXCEL_SHEET_NAME}")
        return

    ws = wb[EXCEL_SHEET_NAME]

    while ws.max_row < EXCEL_INSERT_ROW:
        ws.append([])

    data = [
        fields["Product Description"],
        fields["Batch/Lot No."],
        fields["Date"],
        fields["SKU"],
        fields["Qty"]
    ]

    for col, val in enumerate(data, start=1):
        ws.cell(row=EXCEL_INSERT_ROW, column=col, value=val)

    wb.save(temp_path)
    shutil.move(temp_path, EXCEL_PATH)
    st.success("💾 Saved to Excel!")


# ===============================================================
# SIDEBAR NAVIGATION
# ===============================================================
st.sidebar.title("💜 Tarte SKU System")

page = st.sidebar.radio(
    "Navigate",
    ["Home", "Upload", "Camera", "Chatbot", "Database", "Excel Tools"]
)

if st.sidebar.button("📂 Open Excel File"):
    open_excel_file()


# ===============================================================
# HOME PAGE
# ===============================================================
if page == "Home":

    st.markdown("""
        <div class="tarte-header">
            <h2>Tarte AI-Powered SKU Intake System</h2>
            <p>Upload → Extract → Parse → Save to Excel + Database</p>
        </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="tarte-card">', unsafe_allow_html=True)
    st.subheader("✨ Welcome")

    st.write("""
    This tool was designed to solve a real workflow challenge observed at Tarte:  
    **managing vendor retain documents efficiently and accurately.**

    Features:
    - Extract SKU data from **DOCX or images**
    - **Camera capture** direct from device
    - Auto-dedupe SQLite database  
    - Auto-insert into **Master Sheet - 12th Floor**, row 757  
    - Bulk batch processing  
    - Search + chatbot interface  
    - Beautiful Tarte-branded UI (purple gradient, gold buttons)
    """)
    st.markdown('</div>', unsafe_allow_html=True)


# ===============================================================
# UPLOAD PAGE
# ===============================================================
elif page == "Upload":

    st.markdown('<div class="tarte-card">', unsafe_allow_html=True)
    st.header("📄 Upload Vendor Documents")
    st.write("Upload one or more DOCX or image files.")

    files = st.file_uploader(
        "Upload Documents",
        type=["docx", "jpg", "jpeg", "png"],
        accept_multiple_files=True
    )

    all_fields = []

    if files:
        st.success(f"{len(files)} files uploaded.")

        for idx, file in enumerate(files, start=1):
            st.markdown("---")
            st.subheader(f"📄 Document {idx}: {file.name}")

            text = extract_text_from_docx(file) if file.name.endswith(".docx") else extract_text_from_image(file)

            with st.expander("🔎 Extracted Text"):
                st.text(text)

            fields = parse_vendor_doc(text)
            all_fields.append(fields)

            st.json(fields)

            if st.button(f"💾 Save {file.name}", key=f"save_{idx}"):
                save_to_db(fields)
                save_to_excel(fields)

        if st.button("🔥 Process ALL"):
            for f in all_fields:
                save_to_db(f)
                save_to_excel(f)
            st.success("All documents processed.")
    
    st.markdown('</div>', unsafe_allow_html=True)


# ===============================================================
# CAMERA CAPTURE
# ===============================================================
elif page == "Camera":

    st.markdown('<div class="tarte-card">', unsafe_allow_html=True)
    st.header("📸 Capture with Camera")

    photo = st.camera_input("Take a photo")

    if photo:
        st.success("📸 Photo captured.")
        text = extract_text_from_image(photo)

        with st.expander("🔎 Extracted Text"):
            st.text(text)

        fields = parse_vendor_doc(text)
        st.json(fields)

        if st.button("💾 Save Photo"):
            save_to_db(fields)
            save_to_excel(fields)

    st.markdown('</div>', unsafe_allow_html=True)


# ===============================================================
# CHATBOT PAGE
# ===============================================================
elif page == "Chatbot":

    st.markdown('<div class="tarte-card">', unsafe_allow_html=True)
    st.header("🤖 Chatbot Mode")

    st.write("""
        Commands:  
        - **process latest**  
        - **process folder**  
        - **search SN52**  
        - **show database**  
        - **clear database**
    """)

    if "messages" not in st.session_state:
        st.session_state.messages = []

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.write(msg["content"])

    user_input = st.chat_input("Enter command...")

    if user_input:
        st.session_state.messages.append({"role": "user", "content": user_input})

        # PROCESS LATEST
        if "process latest" in user_input.lower():
            folder = "data/specs/"
            files = [os.path.join(folder, f) for f in os.listdir(folder)
                     if f.endswith((".docx", ".jpg", ".jpeg", ".png"))]

            if not files:
                reply = "❌ No documents found."
            else:
                latest = max(files, key=os.path.getmtime)
                text = extract_text_from_docx(latest) if latest.endswith(".docx") else extract_text_from_image(latest)

                fields = parse_vendor_doc(text)
                save_to_db(fields)
                save_to_excel(fields)
                reply = f"Processed latest file: {os.path.basename(latest)}"

        # PROCESS FOLDER
        elif "process folder" in user_input.lower() or "process all" in user_input.lower():
            folder = "data/specs/"
            files = [os.path.join(folder, f) for f in os.listdir(folder)
                     if f.endswith((".docx", ".jpg", ".jpeg", ".png"))]

            if not files:
                reply = "❌ Folder empty."
            else:
                for f in files:
                    text = extract_text_from_docx(f) if f.endswith(".docx") else extract_text_from_image(f)
                    fields = parse_vendor_doc(text)
                    save_to_db(fields)
                reply = f"🔥 Processed {len(files)} documents."

        # SEARCH
        elif "search" in user_input.lower():
            term = user_input.split("search", 1)[1].strip()
            df = search_db(term)
            st.dataframe(df)
            reply = f"Search results for {term}"

        # SHOW DB
        elif "show database" in user_input.lower():
            conn = sqlite3.connect("sku_catalog.db")
            df = pd.read_sql_query("SELECT * FROM sku_catalog", conn)
            conn.close()
            st.dataframe(df)
            reply = "Displaying database"

        # CLEAR DATABASE
        elif "clear database" in user_input.lower():
            clear_database()
            reply = "Database cleared."

        else:
            reply = "Unknown command."

        st.session_state.messages.append({"role": "assistant", "content": reply})

    st.markdown('</div>', unsafe_allow_html=True)


# ===============================================================
# DATABASE PAGE
# ===============================================================
elif page == "Database":

    st.markdown('<div class="tarte-card">', unsafe_allow_html=True)
    st.header("📊 Database Search & Viewer")

    init_db()

    search_term = st.text_input("Search database...")

    if search_term:
        df = search_db(search_term)
        st.write(f"Results for '{search_term}':")
        st.dataframe(df)
    else:
        conn = sqlite3.connect("sku_catalog.db")
        df = pd.read_sql_query("SELECT * FROM sku_catalog", conn)
        conn.close()
        st.write("All entries:")
        st.dataframe(df)

    if st.button("🧹 Clear Database"):
        clear_database()
        st.success("Database cleared.")

    st.markdown('</div>', unsafe_allow_html=True)


# ===============================================================
# EXCEL TOOLS PAGE
# ===============================================================
elif page == "Excel Tools":

    st.markdown('<div class="tarte-card">', unsafe_allow_html=True)
    st.header("📂 Excel Tools")

    st.write(f"Excel Path: `{EXCEL_PATH}`")
    st.write(f"Sheet Name: `{EXCEL_SHEET_NAME}`")
    st.write(f"Row Insert: `{EXCEL_INSERT_ROW}`")

    if st.button("📂 Open Excel File"):
        open_excel_file()

    st.markdown("---")
    st.write("""
        If the sheet does not update:
        - Make sure Excel is **closed**
        - Confirm the sheet name exactly matches
        - Confirm row number is correct
    """)

    st.markdown('</div>', unsafe_allow_html=True)
