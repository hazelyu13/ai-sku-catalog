import streamlit as st
import sqlite3
import pytesseract
from PIL import Image
from docx import Document
import pandas as pd
import openpyxl
import os
import shutil

# ===============================================================
# 1. OCR & DOCX Extraction
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
# 2. Parse Fields
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
        for key in fields:
            if line.lower().startswith(key.lower()):
                parts = line.split(":", 1)
                if len(parts) == 2:
                    fields[key] = parts[1].strip()

    return fields

# ===============================================================
# 3. Database Operations
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

    sku = fields["SKU"]
    batch = fields["Batch/Lot No."]

    cur.execute("SELECT COUNT(*) FROM sku_catalog WHERE sku = ? AND batch_lot = ?", (sku, batch))
    exists = cur.fetchone()[0]

    if exists > 0:
        st.warning(f"⚠️ Entry already exists (SKU: {sku}, Batch: {batch}) — skipped.")
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
        st.success(f"✅ Saved: SKU {sku}, Batch {batch}")

    conn.close()


def search_database(search_term):
    conn = sqlite3.connect("sku_catalog.db")
    df = pd.read_sql_query(
        "SELECT * FROM sku_catalog WHERE sku LIKE ? OR batch_lot LIKE ?",
        conn,
        params=(f"%{search_term}%", f"%{search_term}%"))
    conn.close()
    return df


def clear_database():
    conn = sqlite3.connect("sku_catalog.db")
    conn.execute("DELETE FROM sku_catalog;")
    conn.commit()
    conn.close()

# ===============================================================
# 4. Excel Save (append to next empty row)
# ===============================================================
def save_to_excel(
    fields,
    excel_path="data/specs/sample_specs.xlsx",
    sheet_name="Master Sheet - 12th Floor"
):

    if not os.path.exists(excel_path):
        st.error(f"❌ Excel file not found at: {excel_path}")
        return

    temp_copy = excel_path.replace(".xlsx", "_temp.xlsx")
    shutil.copyfile(excel_path, temp_copy)

    try:
        wb = openpyxl.load_workbook(temp_copy)
    except Exception as e:
        st.error(f"❌ Could not open Excel file.\n\nError: {e}")
        return

    if sheet_name not in wb.sheetnames:
        st.error(f"❌ Sheet '{sheet_name}' not found. Sheets available: {wb.sheetnames}")
        return

    ws = wb[sheet_name]
    target_row = ws.max_row + 1

    data = [
        fields.get("Product Description"),
        fields.get("Batch/Lot No."),
        fields.get("Date"),
        fields.get("SKU"),
        fields.get("Qty"),
    ]

    for col, value in enumerate(data, start=1):
        ws.cell(row=target_row, column=col, value=value)

    wb.save(temp_copy)
    shutil.move(temp_copy, excel_path)
    st.info(f"✅ Added to Excel row {target_row} in '{sheet_name}'")

# ===============================================================
# 5. NEW — BULK PROCESSING FEATURE
# ===============================================================
def process_all_files_in_folder(folder="data/specs/"):
    files = [os.path.join(folder, f) for f in os.listdir(folder)
             if f.endswith((".docx", ".jpg", ".png", ".jpeg"))]

    if not files:
        return "❌ No files found to process."

    processed_count = 0

    for file in files:
        text = extract_text_from_docx(file) if file.endswith(".docx") else extract_text_from_image(file)
        fields = parse_vendor_doc(text)
        save_to_db(fields)
        save_to_excel(fields)
        processed_count += 1

    return f"✅ Bulk import complete — {processed_count} files processed."

# ===============================================================
# 6. Streamlit UI + Chat Agent
# ===============================================================
st.title("🤖 AI SKU Agent — Upload, Search, Store, Bulk Import")

st.markdown("""
✨ Now supports **bulk importing multiple vendor documents at once.**
Use the chatbot or the button below.

**Chat Commands**
- `process` → imports latest doc
- `process all` → bulk import all docs/images in folder
- `search SN52`
- `clear database`
""")

# Chat history
if "messages" not in st.session_state:
    st.session_state.messages = []

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

user_input = st.chat_input("Type a command...")

if user_input:
    st.session_state.messages.append({"role": "user", "content": user_input})
    response = None

    if "process all" in user_input.lower() or "bulk" in user_input.lower():
        response = process_all_files_in_folder()

    elif "process" in user_input.lower():
        folder = "data/specs/"
        files = [os.path.join(folder, f) for f in os.listdir(folder)
                 if f.endswith((".docx", ".jpg", ".png", ".jpeg"))]
        if not files:
            response = "❌ No files found."
        else:
            latest = max(files, key=os.path.getmtime)
            text = extract_text_from_docx(latest) if latest.endswith(".docx") else extract_text_from_image(latest)
            fields = parse_vendor_doc(text)
            save_to_db(fields)
            save_to_excel(fields)
            response = f"✅ Imported latest SKU: {fields.get('SKU')}"

    elif "search" in user_input.lower():
        term = user_input.split()[-1]
        df = search_database(term)
        st.dataframe(df)
        response = f"📊 Showing results for `{term}`"

    elif "clear database" in user_input.lower():
        clear_database()
        response = "🧹 Database cleared."

    else:
        response = "🤖 I can `process`, `process all`, `search`, or `clear database`."

    st.session_state.messages.append({"role": "assistant", "content": response})

# ===============================================================
# Manual Upload Section
# ===============================================================
st.divider()
st.header("📄 Manual Upload")

uploaded_file = st.file_uploader("Upload DOCX or Image", type=["jpg", "png", "jpeg", "docx"])

if uploaded_file:
    text = extract_text_from_docx(uploaded_file) if uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document" else extract_text_from_image(uploaded_file)

    st.subheader("🔎 Extracted Text")
    st.text(text)

    fields = parse_vendor_doc(text)
    st.subheader("📦 Parsed Fields")
    st.json(fields)

    if st.button("💾 Save to Database & Excel"):
        save_to_db(fields)
        save_to_excel(fields)
        st.success("✅ Saved successfully.")

# Bulk processing button in UI
if st.button("⚡ Process ALL documents in /data/specs/"):
    result = process_all_files_in_folder()
    st.success(result)
