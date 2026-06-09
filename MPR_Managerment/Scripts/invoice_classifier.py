"""
Invoice Classifier
- Đọc text từ PDF (PyMuPDF)
- Tìm PONo bằng label trước, sau đó so khớp toàn bộ DB
- Tra INV_Link từ DB theo PONo → ProjectCode → ProjectInfo
- Copy PDF vào đúng thư mục dự án
- Kết quả trả về JSON
"""

import os
import sys
import io
import json
import re
import shutil
import argparse
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

try:
    import fitz  # PyMuPDF
except ImportError:
    print(json.dumps({"success": False, "error": "pymupdf chưa cài. Chạy: pip install pymupdf"}))
    sys.exit(1)

try:
    import pyodbc
except ImportError:
    print(json.dumps({"success": False, "error": "pyodbc chưa cài. Chạy: pip install pyodbc"}))
    sys.exit(1)

# ── Cấu hình kết nối DB ────────────────────────────────────────────────────────
DB_CONN = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=tcp:dlhi.database.windows.net,1433;"
    "DATABASE=MPR_Management;"
    "UID=davidhoang;"
    "PWD=Hoangquyen@1905;"
    "Encrypt=yes;TrustServerCertificate=no;"
)

# Label phổ biến trong PDF hóa đơn để nhận dạng số PO
PO_LABELS = [
    r"(?:our\s+)?p\.?o\.?\s*(?:no|number|#|num)[\s:.\-]*([A-Z0-9\-\/]+)",
    r"purchase\s+order\s*(?:no|number|#)?[\s:.\-]*([A-Z0-9\-\/]+)",
    r"\bpo\s*[:#\-]\s*([A-Z0-9\-\/]+)",
    r"order\s+(?:no|number|ref)[\s:.\-]*([A-Z0-9\-\/]+)",
]

UNCLASSIFIED_FOLDER = "Chưa phân loại"


def extract_text(pdf_path: str) -> str:
    """Đọc toàn bộ text từ PDF."""
    try:
        doc = fitz.open(pdf_path)
        return "\n".join(page.get_text() for page in doc)
    except Exception as e:
        return ""


def find_po_by_label(text: str) -> list[str]:
    """Tìm PONo theo label (PO No, Purchase Order...)."""
    candidates = []
    text_upper = text.upper()
    for pattern in PO_LABELS:
        for m in re.finditer(pattern, text_upper, re.IGNORECASE | re.MULTILINE):
            val = m.group(1).strip().rstrip(".,;")
            if len(val) >= 4:
                candidates.append(val)
    return list(dict.fromkeys(candidates))  # dedup giữ thứ tự


def load_po_list(conn) -> list[dict]:
    """Tải toàn bộ danh sách PONo + INV_Link từ DB."""
    sql = """
        SELECT h.PONo, h.ProjectCode, p.INV_Link, p.ProjectName
        FROM PO_head h
        LEFT JOIN ProjectInfo p ON p.ProjectCode = h.ProjectCode
        WHERE h.PONo IS NOT NULL AND h.PONo != ''
        ORDER BY LEN(h.PONo) DESC
    """
    cursor = conn.cursor()
    cursor.execute(sql)
    rows = []
    for row in cursor.fetchall():
        rows.append({
            "po_no": (row[0] or "").strip(),
            "project_code": (row[1] or "").strip(),
            "inv_link": (row[2] or "").strip(),
            "project_name": (row[3] or "").strip(),
        })
    return rows


def match_po_in_text(text: str, po_list: list[dict]) -> dict | None:
    """So khớp text PDF với danh sách PONo từ DB (longest match first)."""
    text_upper = text.upper()
    for po in po_list:
        po_no = po["po_no"].upper()
        if not po_no:
            continue
        # Tìm exact match với word boundary
        pattern = r'(?<![A-Z0-9\-])' + re.escape(po_no) + r'(?![A-Z0-9\-])'
        if re.search(pattern, text_upper):
            return po
    return None


def detect_po(pdf_path: str, po_list: list[dict]) -> dict | None:
    """Đọc PDF và trả về PO match (hoặc None). Không copy file."""
    text = extract_text(pdf_path)
    if not text.strip():
        return None

    # Bước 1: tìm theo label
    for candidate in find_po_by_label(text):
        for po in po_list:
            if candidate.upper() == po["po_no"].upper():
                return po

    # Bước 2: so khớp toàn bộ text
    return match_po_in_text(text, po_list)


def pick_best_file(paths: list[str]) -> str:
    """Chọn file tốt nhất trong nhóm trùng PO: ưu tiên file lớn nhất (đủ trang)."""
    return max(paths, key=lambda p: os.path.getsize(p))


def sanitize_po_for_filename(po_no: str) -> str:
    """Loại ký tự không hợp lệ trong tên file Windows."""
    return re.sub(r'[\\/:*?"<>|]', "_", po_no)


def classify_all(pdf_files: list[str], unclassified_base: str, conn) -> list[dict]:
    """
    Phân loại toàn bộ danh sách PDF với dedup theo PO:
    - Các file cùng PO: chỉ giữ 1 (file lớn nhất), copy với tên INV_<PONo>.pdf, xóa các file còn lại.
    - File không tìm được PO: copy vào thư mục Chưa phân loại.
    """
    po_list = load_po_list(conn)

    # Giai đoạn 1: detect PO cho từng file
    po_groups: dict[str, list[str]] = {}   # po_no → [pdf_paths]
    no_po_files: list[str] = []
    po_meta: dict[str, dict] = {}          # po_no → matched_po dict

    for pdf_path in pdf_files:
        matched = detect_po(pdf_path, po_list)
        if matched:
            key = matched["po_no"].upper()
            po_groups.setdefault(key, []).append(pdf_path)
            po_meta[key] = matched
        else:
            no_po_files.append(pdf_path)

    results = []

    # Giai đoạn 2: xử lý từng nhóm PO
    for po_key, group_files in po_groups.items():
        meta = po_meta[po_key]
        po_no = meta["po_no"]
        safe_po = sanitize_po_for_filename(po_no)
        dest_filename = f"INV_{safe_po}.pdf"

        best_file = pick_best_file(group_files)
        duplicates = [f for f in group_files if f != best_file]

        if meta["inv_link"]:
            dest_dir = meta["inv_link"]
            os.makedirs(dest_dir, exist_ok=True)
            dest_path = os.path.join(dest_dir, dest_filename)
            shutil.copy2(best_file, dest_path)
            status = "classified"
        else:
            unclassified = os.path.join(unclassified_base, "Thiếu INV_Link")
            os.makedirs(unclassified, exist_ok=True)
            dest_path = os.path.join(unclassified, dest_filename)
            shutil.copy2(best_file, dest_path)
            status = "no_inv_link"

        # Xóa tất cả file gốc trong nhóm (kể cả best_file sau khi đã copy)
        for f in group_files:
            try:
                os.remove(f)
            except Exception:
                pass

        results.append({
            "file": os.path.basename(best_file),
            "status": status,
            "po_no": po_no,
            "project_code": meta["project_code"],
            "project_name": meta["project_name"],
            "dest": dest_path,
            "duplicates_removed": len(duplicates),
            "duplicates": [os.path.basename(d) for d in duplicates],
            **({"reason": "Dự án chưa có INV_Link trong database"} if status == "no_inv_link" else {})
        })

    # Giai đoạn 3: file không tìm được PO → Chưa phân loại
    unclassified_dir = os.path.join(unclassified_base, UNCLASSIFIED_FOLDER)
    os.makedirs(unclassified_dir, exist_ok=True)
    for pdf_path in no_po_files:
        filename = os.path.basename(pdf_path)
        dest = os.path.join(unclassified_dir, filename)
        if os.path.exists(dest):
            base, ext = os.path.splitext(filename)
            dest = os.path.join(unclassified_dir, f"{base}_{int(os.path.getmtime(pdf_path))}{ext}")
        shutil.copy2(pdf_path, dest)
        try:
            os.remove(pdf_path)
        except Exception:
            pass
        results.append({
            "file": filename,
            "status": "unclassified",
            "reason": "Không tìm thấy PO khớp trong database",
            "dest": dest
        })

    return results


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--files", nargs="+", help="Danh sách file PDF cần phân loại")
    parser.add_argument("--scan-dir", help="Quét tất cả PDF trong thư mục này")
    parser.add_argument("--unclassified-base", required=True, help="Thư mục base chứa 'Chưa phân loại'")
    args = parser.parse_args()

    pdf_files = []
    if args.files:
        pdf_files = args.files
    elif args.scan_dir:
        pdf_files = [
            str(p) for p in Path(args.scan_dir).rglob("*.pdf")
            if UNCLASSIFIED_FOLDER not in str(p) and "Thiếu INV_Link" not in str(p)
        ]

    if not pdf_files:
        print(json.dumps({"success": True, "results": [], "summary": {"total": 0}}))
        return

    try:
        conn = pyodbc.connect(DB_CONN, timeout=15)
    except Exception as e:
        print(json.dumps({"success": False, "error": f"Không kết nối được DB: {e}"}))
        sys.exit(1)

    results = classify_all(pdf_files, args.unclassified_base, conn)
    conn.close()

    classified = sum(1 for r in results if r["status"] == "classified")
    print(json.dumps({
        "success": True,
        "results": results,
        "summary": {
            "total": len(results),
            "classified": classified,
            "unclassified": len(results) - classified
        }
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
