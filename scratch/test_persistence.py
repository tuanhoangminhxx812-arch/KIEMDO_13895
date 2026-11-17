import os
import sys
import json
import datetime
import io

APP_DIR = os.path.abspath(".")
SAVED_FILE_PATH = os.path.join(APP_DIR, "saved_gl094.xlsx")
DEFAULT_FILE_PATH = os.path.join(APP_DIR, "GL_094_TK13895 (010126-311026).xlsx")
META_FILE_PATH = os.path.join(APP_DIR, "saved_gl094_meta.json")

print("Checking DEFAULT_FILE_PATH exists:", os.path.exists(DEFAULT_FILE_PATH))

def load_saved_meta():
    if os.path.exists(META_FILE_PATH):
        try:
            with open(META_FILE_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return None

def save_active_file(file_bytes, filename):
    with open(SAVED_FILE_PATH, "wb") as f:
        f.write(file_bytes)
    meta = {
        "filename": filename,
        "updated_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "size_kb": round(len(file_bytes) / 1024, 1)
    }
    with open(META_FILE_PATH, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    return meta

# Test initial auto-load
active_bytes = None
active_filename = None
active_updated_at = None

if os.path.exists(SAVED_FILE_PATH):
    with open(SAVED_FILE_PATH, "rb") as f:
        active_bytes = f.read()
    meta = load_saved_meta()
    active_filename = meta.get("filename", "saved_gl094.xlsx") if meta else "saved_gl094.xlsx"
    active_updated_at = meta.get("updated_at", "Đã lưu trước đó") if meta else "Đã lưu trước đó"
else:
    if os.path.exists(DEFAULT_FILE_PATH):
        with open(DEFAULT_FILE_PATH, "rb") as f:
            active_bytes = f.read()
        active_filename = "GL_094_TK13895 (010126-311026).xlsx"
        active_updated_at = "Bản gốc mặc định (01/01 - 31/10/2026)"
        save_active_file(active_bytes, active_filename)

print("Active filename:", active_filename)
print("Active updated_at:", active_updated_at)
print("Active bytes length:", len(active_bytes) if active_bytes else 0)
print("Saved file exists now:", os.path.exists(SAVED_FILE_PATH))
print("Meta file exists now:", os.path.exists(META_FILE_PATH))
print("Meta content:", load_saved_meta())
