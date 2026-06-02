"""
Outlook Invoice Attachment Downloader
Dùng COM Automation để kết nối Outlook Desktop đang chạy,
lọc email theo rules và tải file PDF đính kèm.
"""

import os
import sys
import io
import json
import re

# Ép stdout dùng UTF-8 để tránh lỗi encode tiếng Việt trên Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import argparse
from datetime import datetime

try:
    import win32com.client
except ImportError:
    print(json.dumps({"success": False, "error": "pywin32 chưa được cài. Chạy: pip install pywin32"}))
    sys.exit(1)


ALLOWED_EXTENSIONS = [".pdf"]
OUTLOOK_FOLDER_NAME = "Invoice"


def sanitize_folder_name(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip()


def get_save_path(base_dir: str, subject: str, sender: str) -> str:
    folder_name = sanitize_folder_name(subject[:60]) if subject else sanitize_folder_name(sender)
    path = os.path.join(base_dir, folder_name)
    os.makedirs(path, exist_ok=True)
    return path


def find_invoice_folder(namespace, folder_name: str):
    """Tìm folder 'Invoice' trong tất cả mailbox."""
    for store in namespace.Stores:
        try:
            root = store.GetRootFolder()
            for folder in root.Folders:
                if folder.Name.lower() == folder_name.lower():
                    return folder
                # Tìm thêm 1 cấp con
                try:
                    for sub in folder.Folders:
                        if sub.Name.lower() == folder_name.lower():
                            return sub
                except Exception:
                    pass
        except Exception:
            pass
    return None


def download_attachments(base_dir: str, days_back: int = 30, mark_processed: bool = False):
    results = []
    errors = []

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        return {"success": False, "error": f"Không thể kết nối Outlook: {e}", "files": [], "errors": []}

    folder = find_invoice_folder(namespace, OUTLOOK_FOLDER_NAME)
    if folder is None:
        return {
            "success": False,
            "error": f"Không tìm thấy folder '{OUTLOOK_FOLDER_NAME}' trong Outlook. Hãy tạo folder này.",
            "files": [],
            "errors": []
        }

    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # Mới nhất trước

    cutoff = datetime.now().timestamp() - (days_back * 86400)

    processed_count = 0
    for msg in messages:
        try:
            # Kiểm tra loại item (MailItem = 43)
            if msg.Class != 43:
                continue

            received = msg.ReceivedTime
            # win32com trả về datetime có timezone, convert để so sánh
            try:
                ts = received.timestamp()
            except Exception:
                import time
                ts = time.time()

            if ts < cutoff:
                break  # Items đã sort theo thời gian, có thể dừng sớm

            subject = msg.Subject or ""
            sender = msg.SenderName or msg.SenderEmailAddress or "Unknown"
            save_path = get_save_path(base_dir, subject, sender)

            for att in msg.Attachments:
                att_name = att.FileName
                ext = os.path.splitext(att_name)[1].lower()
                if ext not in ALLOWED_EXTENSIONS:
                    continue

                # Tránh ghi đè: thêm timestamp nếu file đã tồn tại
                dest = os.path.join(save_path, att_name)
                if os.path.exists(dest):
                    base, ex = os.path.splitext(att_name)
                    ts_str = datetime.now().strftime("%Y%m%d%H%M%S")
                    dest = os.path.join(save_path, f"{base}_{ts_str}{ex}")

                att.SaveAsFile(dest)
                results.append({
                    "file": dest,
                    "subject": subject,
                    "sender": sender,
                    "received": str(msg.ReceivedTime)
                })

            processed_count += 1

        except Exception as e:
            errors.append(str(e))
            continue

    return {
        "success": True,
        "emails_processed": processed_count,
        "files_downloaded": len(results),
        "files": results,
        "errors": errors
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--save-dir", required=True, help="Thư mục gốc để lưu file")
    parser.add_argument("--days-back", type=int, default=30, help="Số ngày nhìn lại (mặc định 30)")
    args = parser.parse_args()

    result = download_attachments(
        base_dir=args.save_dir,
        days_back=args.days_back
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
