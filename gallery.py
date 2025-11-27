import os
import re


def save_image_to_gallery(gallery_folder: str, fields: dict, file_obj, original_name: str):
    os.makedirs(gallery_folder, exist_ok=True)

    sku = fields.get("SKU") or "NOSKU"
    lot = fields.get("Batch/Lot No.") or "NOLOT"

    sku_safe = re.sub(r"[^A-Za-z0-9\-]", "", str(sku))
    lot_safe = re.sub(r"[^A-Za-z0-9\-]", "", str(lot))

    ext = os.path.splitext(original_name)[1].lower()
    if ext not in [".jpg", ".jpeg", ".png"]:
        ext = ".jpg"

    filename = f"SKU_{sku_safe}_LOT_{lot_safe}{ext}"
    path = os.path.join(gallery_folder, filename)

    with open(path, "wb") as out:
        out.write(file_obj.getbuffer())

    return path


def list_gallery_images(gallery_folder: str):
    if not os.path.exists(gallery_folder):
        os.makedirs(gallery_folder)
        return []
    return [
        os.path.join(gallery_folder, img)
        for img in os.listdir(gallery_folder)
        if img.lower().endswith((".jpg", ".jpeg", ".png"))
    ]


def build_gallery_lookup(db_df):
    lookup = {}
    if db_df is None or db_df.empty:
        return lookup
    for _, row in db_df.iterrows():
        key = (str(row["sku"]), str(row["batch_lot"]))
        lookup[key] = row
    return lookup
