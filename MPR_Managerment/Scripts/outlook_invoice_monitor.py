"""
Outlook Invoice Monitor — chạy nền, lắng nghe email mới qua COM event.
Ghi kết quả ra file JSON log để C# app đọc.
Dùng: python outlook_invoice_monitor.py --save-dir "D:\..." --log-file "D:\...\monitor.log"
"""

import os
import sys
import io
import json
import re
import time
import argparse
import threading
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

try:
    import win32com.client
    import pythoncom
except ImportError:
    print(json.dumps({"event": "error", "message": "pywin32 chưa được cài. Chạy: pip install pywin32"}), flush=True)
    sys.exit(1)

ALLOWED_EXTENSIONS = [".pdf"]
OUTLOOK_FOLDER_NAME = "Invoice"


def sanitize_folder_name(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip()


def process_mail_item(entry_id, namespace, base_dir: str, log_file: str):
    """Xử lý một email khi có sự kiện NewMailEx."""
    try:
        msg = namespace.GetItemFromID(entry_id)
        if msg.Class != 43:
            return

        subject = msg.Subject or ""
        sender = msg.SenderName or msg.SenderEmailAddress or "Unknown"
        folder_name = sanitize_folder_name(subject[:60]) if subject else sanitize_folder_name(sender)
        save_path = os.path.join(base_dir, folder_name)
        os.makedirs(save_path, exist_ok=True)

        downloaded = []
        for att in msg.Attachments:
            att_name = att.FileName
            ext = os.path.splitext(att_name)[1].lower()
            if ext not in ALLOWED_EXTENSIONS:
                continue

            dest = os.path.join(save_path, att_name)
            if os.path.exists(dest):
                base, ex = os.path.splitext(att_name)
                ts_str = datetime.now().strftime("%Y%m%d%H%M%S")
                dest = os.path.join(save_path, f"{base}_{ts_str}{ex}")

            att.SaveAsFile(dest)
            downloaded.append(dest)

        if downloaded:
            event = {
                "event": "downloaded",
                "time": datetime.now().isoformat(),
                "subject": subject,
                "sender": sender,
                "files": downloaded
            }
            print(json.dumps(event, ensure_ascii=False), flush=True)
            _append_log(log_file, event)

    except Exception as e:
        err = {"event": "error", "time": datetime.now().isoformat(), "message": str(e)}
        print(json.dumps(err, ensure_ascii=False), flush=True)
        _append_log(log_file, err)


def _append_log(log_file: str, entry: dict):
    try:
        logs = []
        if os.path.exists(log_file):
            with open(log_file, "r", encoding="utf-8") as f:
                logs = json.load(f)
        logs.append(entry)
        # Giữ tối đa 500 dòng log
        if len(logs) > 500:
            logs = logs[-500:]
        with open(log_file, "w", encoding="utf-8") as f:
            json.dump(logs, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


class OutlookHandler:
    def __init__(self, namespace, base_dir: str, log_file: str):
        self.namespace = namespace
        self.base_dir = base_dir
        self.log_file = log_file

    def OnNewMailEx(self, entry_ids: str):
        for entry_id in entry_ids.split(","):
            entry_id = entry_id.strip()
            if entry_id:
                # Xử lý trong thread riêng để không block event loop
                threading.Thread(
                    target=process_mail_item,
                    args=(entry_id, self.namespace, self.base_dir, self.log_file),
                    daemon=True
                ).start()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--save-dir", required=True)
    parser.add_argument("--log-file", required=True)
    parser.add_argument("--pid-file", default="")
    args = parser.parse_args()

    os.makedirs(args.save_dir, exist_ok=True)

    # Ghi PID để C# có thể kill khi cần
    if args.pid_file:
        with open(args.pid_file, "w") as f:
            f.write(str(os.getpid()))

    pythoncom.CoInitialize()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(json.dumps({"event": "error", "message": f"Không thể kết nối Outlook: {e}"}), flush=True)
        sys.exit(1)

    handler = win32com.client.WithEvents(outlook, OutlookHandler)
    handler.namespace = namespace
    handler.base_dir = args.save_dir
    handler.log_file = args.log_file

    print(json.dumps({
        "event": "started",
        "time": datetime.now().isoformat(),
        "pid": os.getpid(),
        "save_dir": args.save_dir
    }, ensure_ascii=False), flush=True)

    _append_log(args.log_file, {
        "event": "started",
        "time": datetime.now().isoformat(),
        "pid": os.getpid()
    })

    # Vòng lặp COM message pump — giữ process sống và xử lý event
    try:
        while True:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.5)
    except KeyboardInterrupt:
        pass
    finally:
        if args.pid_file and os.path.exists(args.pid_file):
            os.remove(args.pid_file)
        print(json.dumps({"event": "stopped", "time": datetime.now().isoformat()}), flush=True)


if __name__ == "__main__":
    main()
