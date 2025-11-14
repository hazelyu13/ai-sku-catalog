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
# 0. OPEN EXCEL FILE
# ===============================================================
def open_excel_file(excel_path="data/specs/sample_specs.xlsx"):
    """Open Excel file in the system default app."""
    try:
        system = platform.system()

        if system == "Darwin":  # macOS
            subprocess.call(["open", excel_path])
        elif system == "Windows":
            os.startfile(excel_path)
        else:  # Linux
            subprocess.call(["xdg-open", excel_path])

        st.success("📂 Excel file opened!")
    except Exception as e:
        st.error(f"❌ Could not open Excel file.\n\nError: {e}")


# ===============================================================
# 1. OCR + DOCX EXTRACTION
# ===============================================================
def extract_text_from_image(image_file):
    img = Image.open(image_file)
    return pytesseract.image_to_string(img)

def extract_text_from_docx(docx_file):
    doc = Document(docx_file)
    text_parts = []

    for para in doc.paragraphs:
        if para.text.strip():
            text_parts.append(para.text.strip())

    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if len(cells) == 2:
                text_parts.append(f"{cells[0]} {cells[1]}")
            elif cells:
                text_parts.append(" | ".join(cells))
    return "\n".join(text_parts)


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
        line = line.strip()
        for key in fields.keys():
            if line.lower().startswith(key.lower()):
                parts = line.split(":", 1)
                if len(parts) == 2:
                    fields[key] = parts[1].strip()
    return fields


# ===============================================================
# 3. DATABASE (dedupe + search)
# ===============================================================
def save_to_db(fields):
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

    cur.execute("SELECT COUNT(*) FROM sku_catalog WHERE sku = ? AND batch_lot = ?",
                (fields["SKU"], fields["Batch/Lot No."]))
    exists = cur.fetchone()[0]

    if exists > 0:
        st.warning(f"⚠️ Duplicate found: SKU {fields['SKU']} Batch {fields['Batch/Lot No.']}")
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
        WHERE product_desc LIKE '%{keyword}%'
        OR sku LIKE '%{keyword}%'
        OR batch_lot LIKE '%{keyword}%'
    """, conn)
    conn.close()
    return df

def clear_database():
    conn = sqlite3.connect("sku_catalog.db")
    cur = conn.cursor()
    cur.execute("DELETE FROM sku_catalog;")
    conn.commit()
    conn.close()


# ===============================================================
# 4. SAVE TO EXCEL (tab + row)
# ===============================================================
def save_to_excel(fields, excel_path="data/specs/sample_specs.xlsx",
                  sheet_name="Master Sheet - 12th Floor", row=757):

    if not os.path.exists(excel_path):
        st.error(f"❌ Excel not found: {excel_path}")
        return

    temp = excel_path.replace(".xlsx", "_temp.xlsx")
    shutil.copyfile(excel_path, temp)

    wb = openpyxl.load_workbook(temp)
    if sheet_name not in wb.sheetnames:
        st.error(f"❌ Sheet '{sheet_name}' not found.")
        return

    ws = wb[sheet_name]

    while ws.max_row < row:
        ws.append([])

    data = [
        fields["Product Description"],
        fields["Batch/Lot No."],
        fields["Date"],
        fields["SKU"],
        fields["Qty"]
    ]

    for col, value in enumerate(data, start=1):
        ws.cell(row=row, column=col, value=value)

    wb.save(temp)
    shutil.move(temp, excel_path)
    st.success(f"✅ Added to Excel at row {row}")


# ===============================================================
# 5. STREAMLIT UI (CHAT + CAMERA + UPLOAD + FOLDER)
# ===============================================================
st.title("📦 AI-Powered Vendor Document → Excel + Database")

# Chat memory
if "messages" not in st.session_state:
    st.session_state.messages = []

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])

# Chat input
user_input = st.chat_input("Commands: process latest | search SN52 | process folder")

if user_input:
    st.session_state.messages.append({"role": "user", "content": user_input})

    # Process latest file
    if "process latest" in user_input.lower():
        folder = "data/specs/"
        files = [os.path.join(folder, f) for f in os.listdir(folder)
                 if f.endswith((".jpg", ".png", ".jpeg", ".docx"))]
        if not files:
            reply = "❌ No files found."
        else:
            latest = max(files, key=os.path.getmtime)
            text = extract_text_from_docx(latest) if latest.endswith(".docx") else extract_text_from_image(latest)
            fields = parse_vendor_doc(text)
            save_to_db(fields)
            save_to_excel(fields)
            reply = f"✅ Processed: {os.path.basename(latest)}"

    elif "process folder" in user_input.lower() or "process all" in user_input.lower():
        folder = "data/specs/"
        files = [os.path.join(folder, f) for f in os.listdir(folder)
                 if f.endswith((".jpg", ".png", ".jpeg", ".docx"))]

        if not files:
            reply = "❌ No files found."
        else:
            for f in files:
                text = extract_text_from_docx(f) if f.endswith(".docx") else extract_text_from_image(f)
                save_to_db(parse_vendor_doc(text))
            reply = f"🔥 Processed {len(files)} files."

    elif "search" in user_input.lower():
        keyword = user_input.split("search", 1)[1].strip()
        df = search_db(keyword)
        st.dataframe(df)
        reply = f"🔎 Results for '{keyword}'"

    elif "show database" in user_input.lower():
        conn = sqlite3.connect("sku_catalog.db")
        df = pd.read_sql_query("SELECT * FROM sku_catalog", conn)
        conn.close()
        st.dataframe(df)
        reply = "📊 Showing database"

    elif "clear database" in user_input.lower():
        clear_database()
        reply = "🧹 Database cleared"

    else:
        reply = "🤖 Commands: process latest | process folder | search <keyword> | show database | clear database"

    st.session_state.messages.append({"role": "assistant", "content": reply})


# ===============================================================
# CAMERA MODE
# ===============================================================
st.divider()
st.subheader("📸 Capture Vendor Document")

if "camera_open" not in st.session_state:
    st.session_state.camera_open = False

if st.button("📷 Open Camera"):
    st.session_state.camera_open = True

camera_file = st.camera_input("Take a picture") if st.session_state.camera_open else None

if camera_file:
    st.success("📸 Photo captured — click Save to process.")
    text = extract_text_from_image(camera_file)
    fields = parse_vendor_doc(text)
    st.json(fields)

    if st.button("💾 Save Camera Image"):
        save_to_db(fields)
        save_to_excel(fields)


# ===============================================================
# MULTI-FILE UPLOAD + PROCESS ALL
# ===============================================================
st.divider()
st.subheader("📄 Upload Multiple Vendor Documents")

uploaded_files = st.file_uploader(
    "Upload one or multiple vendor documents",
    type=["jpg", "png", "jpeg", "docx"],
    accept_multiple_files=True
)

all_fields = []

if uploaded_files:
    st.success(f"📁 {len(uploaded_files)} files uploaded.")

    for idx, file in enumerate(uploaded_files, start=1):
        st.write(f"### Document {idx}: {file.name}")

        text = extract_text_from_docx(file) if file.name.endswith(".docx") else extract_text_from_image(file)
        st.text_area(f"Extracted Text ({file.name})", text, height=150)

        fields = parse_vendor_doc(text)
        all_fields.append(fields)
        st.json(fields)

        if st.button(f"💾 Save {file.name}", key=f"save_{idx}"):
            save_to_db(fields)
            save_to_excel(fields)

    if st.button("🔥 Process ALL Uploaded Documents"):
        for f in all_fields:
            save_to_db(f)
            save_to_excel(f)
        st.success("✅ All documents processed!")


# ===============================================================
# OPEN EXCEL FILE BUTTON
# ===============================================================
st.divider()
st.subheader("📂 Open Excel File")

if st.button("Open Master Excel Sheet"):
    open_excel_file("data/specs/sample_specs.xlsx")

