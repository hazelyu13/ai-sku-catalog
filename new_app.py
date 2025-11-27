# ===============================================================
#  TARTE BEAUTY  ·  LUXURY DASHBOARD
#  AI-Powered SKU Intake System (Refactored)
# ===============================================================

import os
import streamlit as st
import pandas as pd

from config import (
    EXCEL_PATH,
    EXCEL_SHEET_NAME,
    EXCEL_INSERT_ROW,
    DB_PATH,
    GALLERY_FOLDER,
    SPECS_FOLDER,
    LOGO_PATH,
)
from utils import open_excel_file, load_image_base64, open_folder
from extractors import (
    extract_text_from_image,
    extract_text_from_docx,
    extract_barcodes_from_image,
)
from parser import parse_vendor_doc, predict_missing_fields, normalize_fields_for_db
from database import (
    init_db,
    save_to_db,
    check_duplicate,
    search_db,
    clear_database,
    export_csv,
    export_excel,
    undo_last_save,
    database_stats,
    get_all_rows,
    get_latest_row,
    get_duplicates_df,
    delete_by_sku,
)
from excel_tools import save_to_excel
from gallery import save_image_to_gallery, list_gallery_images, build_gallery_lookup

# ===============================================================
# STREAMLIT CONFIG
# ===============================================================
st.set_page_config(
    page_title="Tarte AI SKU System",
    page_icon="💜",
    layout="wide",
)

# Load CSS theme
with open("styles/theme.css") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Ensure DB exists
init_db(DB_PATH)

# ===============================================================
# SIDEBAR NAV + LOGO
# ===============================================================
logo_base64 = None
try:
    logo_base64 = load_image_base64(LOGO_PATH)
except FileNotFoundError:
    logo_base64 = None

with st.sidebar:
    if logo_base64:
        st.markdown(
            f"""
            <div style="text-align:center; margin-bottom:10px;">
                <img src="data:image/png;base64,{logo_base64}"
                     style="width:140px; margin-top:-10px;">
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.title("Tarte SKU System")

    page = st.radio(
        "",
        [
            "Home",
            "Upload",
            "Camera",
            "Chatbot",
            "Database",
            "Gallery",
            "Stats",
            "Settings",
            "Excel Tools",
        ],
        label_visibility="collapsed",
    )

    if st.button("📂 Open Excel File"):
        open_excel_file(EXCEL_PATH)

    if st.button("🖼 Open Gallery Folder"):
        open_folder(GALLERY_FOLDER)

    if st.button("🔄 Reset App"):
        st.session_state.clear()
        st.experimental_rerun()

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
# PAGE HELPERS
# ===============================================================


def render_file_badge(file_name: str):
    if file_name.lower().endswith(".docx"):
        badge = "DOCX"
    else:
        badge = "IMG"
    st.markdown(
        f'<span class="tarte-badge">{badge}</span> <span class="tarte-file-name">{file_name}</span>',
        unsafe_allow_html=True,
    )


def render_duplicate_badge():
    st.markdown(
        '<span class="dup-badge">Duplicate</span>',
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
• Extract text + barcodes and parse core SKU fields  
• Review and edit before saving  
• Store records in a duplicate-safe SQLite database  
• Sync rows directly into the master Excel tracker  
• Use the chatbot to trigger actions or search data  
        """
    )

# ===============================================================
# UPLOAD PAGE
# ===============================================================
elif page == "Upload":

    st.header("📄 Upload Vendor Documents")
    st.write("Upload one or more DOCX or image files. Review fields before saving.")

    files = st.file_uploader(
        "Upload one or more files",
        type=["docx", "jpg", "jpeg", "png"],
        accept_multiple_files=True,
        help="Upload retain specs, artwork proofs, or vendor docs.",
    )

    if files:
        st.success(f"{len(files)} files uploaded.")

        for idx, f in enumerate(files, start=1):
            st.markdown("---")
            cols_header = st.columns([3, 1])
            with cols_header[0]:
                st.subheader(f"📄 File {idx}")
                render_file_badge(f.name)

            with cols_header[1]:
                st.write("")  # spacing

            # Extract text + barcodes
            if f.name.lower().endswith(".docx"):
                text = extract_text_from_docx(f)
                barcodes = []
                is_image_file = False
            else:
                text = extract_text_from_image(f)
                barcodes = extract_barcodes_from_image(f)
                is_image_file = True

            if barcodes:
                st.info(f"🔍 Detected barcodes / QR codes: {', '.join(barcodes)}")

            fields = parse_vendor_doc(text)
            fields, ai_flags = predict_missing_fields(text, fields)

            # If barcode and no SKU, use barcode
            if barcodes and not fields.get("SKU"):
                fields["SKU"] = barcodes[0]

            # Normalize for duplicate check
            normalized_fields = normalize_fields_for_db(fields)
            duplicate_row = check_duplicate(DB_PATH, normalized_fields)

            if duplicate_row is not None:
                st.warning(
                    f"⚠️ Potential duplicate for SKU **{normalized_fields.get('SKU')}** / "
                    f"Lot **{normalized_fields.get('Batch/Lot No.')}**",
                    icon="🚨",
                )
                render_duplicate_badge()
                dup_df = pd.DataFrame(
                    [duplicate_row],
                    columns=[
                        "id",
                        "product_desc",
                        "batch_lot",
                        "date",
                        "sku",
                        "qty",
                        "sku_norm",
                        "batch_lot_norm",
                    ],
                )
                st.dataframe(dup_df.drop(columns=["sku_norm", "batch_lot_norm"]), use_container_width=True)

            edited_fields = review_fields_ui = \
                __import__("parser").parser.review_fields_ui(
                    fields,
                    text,
                    key_prefix=f"file_{idx}",
                    ai_flags=ai_flags,
                )

            cols_actions = st.columns(3)
            with cols_actions[0]:
                confirm = st.button(
                    f"✅ Confirm & Save {f.name}",
                    key=f"confirm_{idx}",
                )
            with cols_actions[1]:
                cancel = st.button(
                    f"❌ Cancel {f.name}",
                    key=f"cancel_{idx}",
                )
            with cols_actions[2]:
                rotate_left = st.button("⟲ Rotate Left (Image)", key=f"rotl_{idx}")
                rotate_right = st.button("⟳ Rotate Right (Image)", key=f"rotr_{idx}")

            # Image rotation preview (does not persist yet)
            if is_image_file and (rotate_left or rotate_right):
                from utils import rotate_image_filelike

                rotated = rotate_image_filelike(f, -90 if rotate_left else 90)
                st.image(rotated, caption="Rotated preview", use_column_width=True)

            if confirm:
                edited_norm = normalize_fields_for_db(edited_fields)
                duplicate_row_final = check_duplicate(DB_PATH, edited_norm)
                if duplicate_row_final is not None:
                    st.error("🚫 Duplicate detected - saving blocked.")
                    if st.button(
                        "🔓 Override & Force Save",
                        key=f"override_{idx}",
                    ):
                        save_to_db(DB_PATH, edited_norm)
                        save_to_excel(edited_norm)
                        if is_image_file:
                            save_image_to_gallery(
                                GALLERY_FOLDER, edited_norm, f, f.name
                            )
                else:
                    save_to_db(DB_PATH, edited_norm)
                    save_to_excel(edited_norm)
                    if is_image_file:
                        save_image_to_gallery(GALLERY_FOLDER, edited_norm, f, f.name)

            elif cancel:
                st.info(f"⏹ Skipped saving for {f.name}.")

# ===============================================================
# CAMERA PAGE
# ===============================================================
elif page == "Camera":

    st.header("📸 Capture Vendor Document")
    photo = st.camera_input("Take a photo of the retain or vendor sheet")

    if photo:
        text = extract_text_from_image(photo)
        barcodes = extract_barcodes_from_image(photo)

        if barcodes:
            st.info(f"🔍 Detected barcodes / QR codes: {', '.join(barcodes)}")

        fields = parse_vendor_doc(text)
        fields, ai_flags = predict_missing_fields(text, fields)

        if barcodes and not fields.get("SKU"):
            fields["SKU"] = barcodes[0]

        normalized_fields = normalize_fields_for_db(fields)
        duplicate_row = check_duplicate(DB_PATH, normalized_fields)

        if duplicate_row is not None:
            st.warning(
                f"⚠️ Potential duplicate for SKU **{normalized_fields.get('SKU')}** / "
                f"Lot **{normalized_fields.get('Batch/Lot No.')}**",
                icon="🚨",
            )
            render_duplicate_badge()
            dup_df = pd.DataFrame(
                [duplicate_row],
                columns=[
                    "id",
                    "product_desc",
                    "batch_lot",
                    "date",
                    "sku",
                    "qty",
                    "sku_norm",
                    "batch_lot_norm",
                ],
            )
            st.dataframe(dup_df.drop(columns=["sku_norm", "batch_lot_norm"]), use_container_width=True)

        edited_fields = __import__("parser").parser.review_fields_ui(
            fields,
            text,
            key_prefix="camera",
            ai_flags=ai_flags,
        )

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            confirm = st.button("✅ Confirm & Save from camera")
        with col_b:
            cancel = st.button("❌ Cancel (discard photo)")
        with col_c:
            rotate_photo = st.button("⟳ Rotate Preview")

        if rotate_photo:
            from utils import rotate_image_filelike

            rotated = rotate_image_filelike(photo, 90)
            st.image(rotated, caption="Rotated preview", use_column_width=True)

        if confirm:
            edited_norm = normalize_fields_for_db(edited_fields)
            duplicate_row_final = check_duplicate(DB_PATH, edited_norm)
            if duplicate_row_final is not None:
                st.error("🚫 Duplicate detected - saving blocked.")
                if st.button("🔓 Override & Force Save", key="override_camera"):
                    save_to_db(DB_PATH, edited_norm)
                    save_to_excel(edited_norm)
                    save_image_to_gallery(
                        GALLERY_FOLDER, edited_norm, photo, "camera.jpg"
                    )
            else:
                save_to_db(DB_PATH, edited_norm)
                save_to_excel(edited_norm)
                save_image_to_gallery(
                    GALLERY_FOLDER, edited_norm, photo, "camera.jpg"
                )

        elif cancel:
            st.info("⏹ Camera capture discarded.")

# ===============================================================
# CHATBOT PAGE
# ===============================================================
elif page == "Chatbot":

    st.header("🤖 Chatbot Mode")
    st.write(
        "Try commands like: `process latest`, `search SN52`, `show database`, "
        "`clear database`, `export csv`, `export excel`, `undo`, `stats`, "
        "`list duplicates`, `open excel`, `show latest`, `help`."
    )

    if "messages" not in st.session_state:
        st.session_state.messages = []

    for m in st.session_state.messages:
        with st.chat_message(m["role"]):
            st.write(m["content"])

    user_input = st.chat_input("Ask me to process latest, search, or run commands")

    if user_input:
        st.session_state.messages.append({"role": "user", "content": user_input})
        reply = ""
        lower = user_input.lower()

        if "process latest" in lower:
            files = [
                os.path.join(SPECS_FOLDER, f)
                for f in os.listdir(SPECS_FOLDER)
                if f.lower().endswith((".docx", ".jpg", ".jpeg", ".png"))
            ]

            if not files:
                reply = "❌ No files found in data/specs."
            else:
                latest = max(files, key=os.path.getmtime)

                if latest.lower().endswith(".docx"):
                    with open(latest, "rb") as fh:
                        text = extract_text_from_docx(fh)
                    barcodes = []
                    is_image_file = False
                else:
                    with open(latest, "rb") as fh:
                        text = extract_text_from_image(fh)
                        fh.seek(0)
                        barcodes = extract_barcodes_from_image(fh)
                    is_image_file = True

                fields = parse_vendor_doc(text)
                fields, _ = predict_missing_fields(text, fields)

                if barcodes and not fields.get("SKU"):
                    fields["SKU"] = barcodes[0]

                norm = normalize_fields_for_db(fields)
                duplicate_row = check_duplicate(DB_PATH, norm)
                if duplicate_row is not None:
                    reply = (
                        f"⚠️ Duplicate detected for SKU {norm.get('SKU')} / "
                        f"Lot {norm.get('Batch/Lot No.')} - not saved.\n\n"
                        "Use the Upload or Camera page to review and override if needed."
                    )
                else:
                    save_to_db(DB_PATH, norm)
                    save_to_excel(norm)
                    reply = (
                        f"Processed and saved latest file: {os.path.basename(latest)}"
                    )

                    if is_image_file:
                        with open(latest, "rb") as fh:
                            class _Tmp:
                                def __init__(self, b):
                                    self._b = b

                                def getbuffer(self):
                                    return self._b

                            buf = fh.read()
                            tmp_file = _Tmp(buf)
                            save_image_to_gallery(
                                GALLERY_FOLDER, norm, tmp_file, os.path.basename(latest)
                            )

        elif "search" in lower:
            term = user_input.split("search", 1)[1].strip()
            df = search_db(DB_PATH, term)
            st.dataframe(df)
            reply = f"Search results for '{term}'."

        elif "show database" in lower:
            df = get_all_rows(DB_PATH)
            st.dataframe(df)
            reply = "Full database view."

        elif "clear database" in lower:
            clear_database(DB_PATH)
            reply = "Database cleared."

        elif "export csv" in lower:
            path = export_csv(DB_PATH)
            reply = f"📤 Exported as CSV: `{path}`"

        elif "export excel" in lower:
            path = export_excel(DB_PATH)
            reply = f"📤 Exported as Excel: `{path}`"

        elif "undo" in lower:
            deleted = undo_last_save(DB_PATH)
            if deleted:
                reply = f"↩️ Undid last save (deleted row ID {deleted})."
            else:
                reply = "Nothing to undo."

        elif "stats" in lower:
            total, duplicates = database_stats(DB_PATH)
            reply = f"📊 Total rows: {total}\n\nDuplicate SKU/Lot combos:\n{duplicates}"

        elif "delete" in lower:
            term = lower.split("delete", 1)[1].strip()
            deleted_count = delete_by_sku(DB_PATH, term.upper())
            reply = f"🗑 Deleted {deleted_count} row(s) with SKU `{term}`."

        elif "list duplicates" in lower:
            df = get_duplicates_df(DB_PATH)
            st.dataframe(df)
            reply = "🔍 Duplicate rows shown above."

        elif "open excel" in lower:
            open_excel_file(EXCEL_PATH)
            reply = "Opening Excel..."

        elif "show latest" in lower:
            df = get_latest_row(DB_PATH)
            st.dataframe(df)
            reply = "📄 Latest saved row shown above."

        elif "help" in lower:
            reply = """
Available Commands:
• process latest  
• search <term>  
• show database  
• clear database  
• export csv  
• export excel  
• delete <SKU>  
• undo  
• stats  
• list duplicates  
• open excel  
• show latest  
• help
"""

        else:
            reply = (
                "Unknown command. Try: `process latest`, `search <term>`, "
                "`show database`, `clear database`, `export csv`, `export excel`, "
                "`undo`, `stats`, `list duplicates`, `open excel`, `show latest`, or `help`."
            )

        st.session_state.messages.append({"role": "assistant", "content": reply})

# ===============================================================
# DATABASE PAGE
# ===============================================================
elif page == "Database":

    st.header("📊 Database Viewer")

    search_term = st.text_input("", placeholder="🔍 Search SKU, lot, or product description")

    if search_term:
        df = search_db(DB_PATH, search_term)
    else:
        df = get_all_rows(DB_PATH)

    st.dataframe(df)

    if st.button("🧹 Clear database"):
        clear_database(DB_PATH)
        st.success("Database cleared.")

# ===============================================================
# GALLERY PAGE
# ===============================================================
elif page == "Gallery":

    st.header("🖼️ Processed Retains Gallery")
    st.write("A visual grid of all saved retains, with filters for SKU, lot, and date.")

    images = list_gallery_images(GALLERY_FOLDER)
    db_df = get_all_rows(DB_PATH)

    if db_df.empty:
        sku_options, lot_options, date_options = [], [], []
    else:
        sku_options = sorted(db_df["sku"].dropna().astype(str).unique().tolist())
        lot_options = sorted(db_df["batch_lot"].dropna().astype(str).unique().tolist())
        date_options = sorted(db_df["date"].dropna().astype(str).unique().tolist())

    filt_col1, filt_col2, filt_col3, filt_col4 = st.columns([1, 1, 1, 2])

    with filt_col1:
        sku_filter = st.selectbox(
            "Filter by SKU",
            ["All"] + sku_options if sku_options else ["All"],
        )

    with filt_col2:
        lot_filter = st.selectbox(
            "Filter by Lot",
            ["All"] + lot_options if lot_options else ["All"],
        )

    with filt_col3:
        date_filter = st.selectbox(
            "Filter by Date",
            ["All"] + date_options if date_options else ["All"],
        )

    with filt_col4:
        search_filter = st.text_input(
            "Search text (SKU, lot, or description)",
            "",
            placeholder="e.g. SN52, holiday set, batch 3392",
        ).strip().lower()

    st.markdown("---")

    if not images:
        st.info("No retains in the gallery yet - save from Upload or Camera first.")
    else:
        cols = st.columns(4)
        lookup = build_gallery_lookup(db_df)

        for idx, img_path in enumerate(sorted(images)):
            basename = os.path.basename(img_path)
            import re

            match = re.match(
                r"SKU_(.+)_LOT_(.+)\.[^.]+$",
                basename,
                re.IGNORECASE,
            )

            display_info = None
            sku_val = None
            lot_val = None

            if match:
                sku_val = match.group(1)
                lot_val = match.group(2)
                display_info = lookup.get((sku_val, lot_val))

            if sku_filter != "All":
                if not sku_val or sku_val != sku_filter:
                    continue

            if lot_filter != "All":
                if not lot_val or lot_val != lot_filter:
                    continue

            if date_filter != "All":
                if display_info is None or str(display_info.get("date")) != date_filter:
                    continue

            if search_filter:
                if display_info is None:
                    continue
                haystack = " ".join(
                    [
                        str(display_info.get("sku", "")).lower(),
                        str(display_info.get("batch_lot", "")).lower(),
                        str(display_info.get("product_desc", "")).lower(),
                    ]
                )
                if search_filter not in haystack:
                    continue

            col = cols[idx % 4]
            with col:
                st.markdown(
                    """
                    <div class="gallery-card">
                    """,
                    unsafe_allow_html=True,
                )

                st.image(img_path, use_column_width=True)

                if display_info is not None:
                    st.markdown(
                        f"""
                        <p class="gallery-meta">
                        <b>SKU:</b> {display_info['sku']}<br>
                        <b>Lot:</b> {display_info['batch_lot']}<br>
                        <b>Date:</b> {display_info['date']}<br>
                        <b>Qty:</b> {display_info['qty']}
                        </p>
                        """,
                        unsafe_allow_html=True,
                    )
                else:
                    st.caption("No matching record found for this image.")

                st.markdown("</div>", unsafe_allow_html=True)

# ===============================================================
# STATS PAGE
# ===============================================================
elif page == "Stats":
    st.header("📈 Retains Stats")

    total_rows, dups = database_stats(DB_PATH)
    st.metric("Total Records", total_rows)

    dup_df = get_duplicates_df(DB_PATH)
    st.subheader("Duplicate SKU / Lot Combos")
    if dup_df.empty:
        st.write("✅ No duplicates found.")
    else:
        st.dataframe(dup_df)

# ===============================================================
# SETTINGS PAGE
# ===============================================================
elif page == "Settings":
    st.header("⚙️ Settings & Info")

    st.write("Configure paths and view app info.")

    st.write("**Excel path:**")
    st.code(EXCEL_PATH)

    st.write("**Excel sheet name:**")
    st.code(EXCEL_SHEET_NAME)

    st.write("**Insert row:**")
    st.code(EXCEL_INSERT_ROW)

    st.write("**Database path:**")
    st.code(DB_PATH)

    st.write("**Gallery folder:**")
    st.code(GALLERY_FOLDER)

    st.write("**Specs folder:**")
    st.code(SPECS_FOLDER)

    st.markdown("---")
    st.write("If paths change in IT/ops, update them in `config.py`.")

# ===============================================================
# EXCEL TOOLS
# ===============================================================
elif page == "Excel Tools":

    st.header("📂 Excel Tools")
    st.write(f"Excel file: `{EXCEL_PATH}`")
    st.write(f"Sheet name: `{EXCEL_SHEET_NAME}`")
    st.write(f"Inserting into row: `{EXCEL_INSERT_ROW}`")

    if st.button("📂 Open Excel file (again)"):
        open_excel_file(EXCEL_PATH)

    st.markdown("---")
    st.write("Tip: close the workbook before running updates from this dashboard.")
