"""
Outlook Invoice Monitor - polling Inbox moi 30 giay, tai PDF dinh kem.
Dung: python outlook_invoice_monitor.py --save-dir "D:/..." --log-file "D:/...monitor.log"
"""

import os
import sys
import io
import json
import re
import time
import argparse
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

try:
    import win32com.client
    import pythoncom
except ImportError:
    print(json.dumps({"event": "error", "message": "pywin32 chua duoc cai. Chay: pip install pywin32"}), flush=True)
    sys.exit(1)

ALLOWED_EXTENSIONS = [".pdf"]
POLL_INTERVAL = 30        # giay giua moi lan scan
INBOX_SCAN_LIMIT = 20     # so email toi da scan moi lan


def sanitize_folder_name(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip()


def _append_log(log_file: str, entry: dict):
    try:
        logs = []
        if os.path.exists(log_file):
            with open(log_file, "r", encoding="utf-8") as f:
                logs = json.load(f)
        logs.append(entry)
        if len(logs) > 500:
            logs = logs[-500:]
        with open(log_file, "w", encoding="utf-8") as f:
            json.dump(logs, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def process_msg(msg, base_dir: str, log_file: str) -> bool:
    """Tai PDF dinh kem cua msg. Tra ve True neu co file duoc tai."""
    try:
        subject = msg.Subject or ""
        sender  = msg.SenderName or msg.SenderEmailAddress or "Unknown"
        folder_name = sanitize_folder_name(subject[:60]) if subject else sanitize_folder_name(sender)
        save_path = os.path.join(base_dir, folder_name)
        os.makedirs(save_path, exist_ok=True)

        downloaded = []
        att_count = 0
        for att in msg.Attachments:
            att_count += 1
            att_name = att.FileName
            ext = os.path.splitext(att_name)[1].lower()
            if ext not in ALLOWED_EXTENSIONS:
                continue
            dest = os.path.join(save_path, att_name)
            if os.path.exists(dest):
                base_n, ex = os.path.splitext(att_name)
                ts_str = datetime.now().strftime("%Y%m%d%H%M%S")
                dest = os.path.join(save_path, f"{base_n}_{ts_str}{ex}")
            att.SaveAsFile(dest)
            downloaded.append(dest)

        if downloaded:
            evt = {
                "event": "downloaded",
                "time": datetime.now().isoformat(),
                "subject": subject,
                "sender": sender,
                "files": downloaded
            }
            print(json.dumps(evt, ensure_ascii=False), flush=True)
            _append_log(log_file, evt)
            return True
        else:
            info = {
                "event": "no_pdf",
                "time": datetime.now().isoformat(),
                "subject": subject,
                "sender": sender,
                "attachments_total": att_count
            }
            print(json.dumps(info, ensure_ascii=False), flush=True)
            _append_log(log_file, info)
            return False
    except Exception as e:
        err = {"event": "error", "time": datetime.now().isoformat(), "message": f"process_msg: {e}"}
        print(json.dumps(err, ensure_ascii=False), flush=True)
        _append_log(log_file, err)
        return False


def get_all_inbox_folders(namespace):
    """Lay tat ca Inbox va Invoice folder tu moi account."""
    folders = []
    try:
        root = namespace.Folders
        for ai in range(1, root.Count + 1):
            store = root[ai]
            try:
                for fi in range(1, store.Folders.Count + 1):
                    folder = store.Folders[fi]
                    name_lower = folder.Name.lower()
                    if name_lower in ("inbox", "invoice", "hop thu den"):
                        folders.append(folder)
            except Exception:
                pass
    except Exception:
        pass
    return folders


def poll_inbox(namespace, base_dir: str, log_file: str, processed_ids: set):
    """Scan tat ca Inbox/Invoice folder, xu ly email chua xu ly."""
    new_count = 0
    folders = get_all_inbox_folders(namespace)
    for folder in folders:
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            for i in range(1, min(items.Count + 1, INBOX_SCAN_LIMIT + 1)):
                try:
                    msg = items[i]
                    if msg.Class != 43:
                        continue
                    entry_id = msg.EntryID
                    if entry_id in processed_ids:
                        continue
                    processed_ids.add(entry_id)
                    if process_msg(msg, base_dir, log_file):
                        new_count += 1
                except Exception as e:
                    err = {"event": "error", "time": datetime.now().isoformat(), "message": f"poll item: {e}"}
                    print(json.dumps(err, ensure_ascii=False), flush=True)
                    _append_log(log_file, err)
        except Exception as e:
            err = {"event": "error", "time": datetime.now().isoformat(), "message": f"poll_folder: {e}"}
            print(json.dumps(err, ensure_ascii=False), flush=True)
            _append_log(log_file, err)
    return new_count


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--save-dir", required=True)
    parser.add_argument("--log-file", required=True)
    parser.add_argument("--pid-file", default="")
    args = parser.parse_args()

    os.makedirs(args.save_dir, exist_ok=True)

    if args.pid_file:
        try:
            os.makedirs(os.path.dirname(args.pid_file), exist_ok=True)
        except Exception:
            pass
        with open(args.pid_file, "w") as f:
            f.write(str(os.getpid()))

    pythoncom.CoInitialize()

    try:
        outlook   = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(json.dumps({"event": "error", "message": f"Khong the ket noi Outlook: {e}"}), flush=True)
        sys.exit(1)

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

    processed_ids: set = set()

    # Scan lan dau de danh dau email cu (khong tai lai)
    try:
        for folder in get_all_inbox_folders(namespace):
            try:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)
                for i in range(1, min(items.Count + 1, INBOX_SCAN_LIMIT + 1)):
                    try:
                        msg = items[i]
                        if msg.Class == 43:
                            processed_ids.add(msg.EntryID)
                    except Exception:
                        pass
            except Exception:
                pass
        info = {"event": "init", "time": datetime.now().isoformat(), "marked_existing": len(processed_ids)}
        print(json.dumps(info, ensure_ascii=False), flush=True)
        _append_log(args.log_file, info)
    except Exception as e:
        _append_log(args.log_file, {"event": "error", "time": datetime.now().isoformat(), "message": f"init scan: {e}"})

    # Polling loop
    last_poll = time.time()
    try:
        while True:
            pythoncom.PumpWaitingMessages()
            now = time.time()
            if now - last_poll >= POLL_INTERVAL:
                poll_info = {"event": "polling", "time": datetime.now().isoformat()}
                print(json.dumps(poll_info, ensure_ascii=False), flush=True)
                _append_log(args.log_file, poll_info)
                poll_inbox(namespace, args.save_dir, args.log_file, processed_ids)
                last_poll = now
            time.sleep(2)
    except KeyboardInterrupt:
        pass
    finally:
        if args.pid_file and os.path.exists(args.pid_file):
            try:
                os.remove(args.pid_file)
            except Exception:
                pass
        print(json.dumps({"event": "stopped", "time": datetime.now().isoformat()}), flush=True)


if __name__ == "__main__":
    main()
