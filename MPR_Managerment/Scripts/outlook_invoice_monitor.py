"""
Outlook Invoice Monitor - polling Inbox moi 30 giay, tai PDF dinh kem,
chuyen tiep email toi nhom ke toan, danh dau da doc, ghi success log.
"""

import os
import sys
import io
import json
import re
import time
import argparse
import subprocess
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
# Inbox quet toi da 50 email moi nhat; Invoice folder quet tat ca (khong gioi han)
INBOX_SCAN_LIMIT = 50

# Khi app C# dong, stdout pipe bi dut — flag cho phep script tiep tuc chay o log-only mode
_stdout_alive = True


# ── Helpers ───────────────────────────────────────────────────────────────────

def _safe_print(obj: dict) -> None:
    global _stdout_alive
    if not _stdout_alive:
        return
    try:
        print(json.dumps(obj, ensure_ascii=False), flush=True)
    except (BrokenPipeError, OSError):
        _stdout_alive = False


def sanitize_folder_name(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip()


def _append_log(log_file: str, entry: dict):
    """Ghi vao monitor_log.json (rolling 500 ban ghi)."""
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


def _append_success_log(success_log_file: str, entry: dict):
    """
    Ghi vao success_log.json — nhat ki vinh vien (KHONG tu xoa).
    Moi ban ghi la mot email da duoc tai file + chuyen tiep thanh cong.
    """
    try:
        os.makedirs(os.path.dirname(success_log_file), exist_ok=True)
        logs = []
        if os.path.exists(success_log_file):
            with open(success_log_file, "r", encoding="utf-8") as f:
                logs = json.load(f)
        logs.append(entry)
        with open(success_log_file, "w", encoding="utf-8") as f:
            json.dump(logs, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ── Core functions ────────────────────────────────────────────────────────────

_TARGET_FOLDER_NAMES = {"inbox", "invoice", "hop thu den"}


def _collect_folders_recursive(parent, results: list, depth: int = 0):
    """Duyet de quy tim tat ca folder co ten khop, toi da 5 cap."""
    if depth > 5:
        return
    try:
        count = parent.Folders.Count
    except Exception:
        return
    for fi in range(1, count + 1):
        try:
            folder = parent.Folders[fi]
            if folder.Name.lower() in _TARGET_FOLDER_NAMES:
                results.append(folder)
            # Van duyet sau de tim "Invoice" long trong "mail 2024/Invoice" v.v.
            _collect_folders_recursive(folder, results, depth + 1)
        except (IndexError, KeyError):
            break
        except Exception:
            pass


def get_all_inbox_folders(namespace):
    """Lay tat ca folder ten khop tu moi account, de quy toan cay thu muc."""
    folders = []
    try:
        root = namespace.Folders
        for ai in range(1, root.Count + 1):
            try:
                store = root[ai]
                _collect_folders_recursive(store, folders)
            except (IndexError, KeyError):
                break
            except Exception:
                pass
    except Exception:
        pass
    return folders


def forward_email(msg, forward_to: list[str]) -> tuple[bool, str]:
    """
    Chuyen tiep email toi danh sach dia chi.
    Tra ve (True, "") neu thanh cong, (False, reason) neu that bai.
    """
    if not forward_to:
        return True, ""
    try:
        fwd = msg.Forward()
        for addr in forward_to:
            addr = addr.strip()
            if addr:
                recipient = fwd.Recipients.Add(addr)
                recipient.Type = 1  # olTo
        fwd.Recipients.ResolveAll()
        fwd.Send()
        return True, ""
    except Exception as e:
        return False, str(e)


def mark_as_read(msg) -> bool:
    """Danh dau email la da doc. Tra ve True neu thanh cong."""
    try:
        msg.UnRead = False
        return True
    except Exception:
        return False


def classify_downloaded_files(downloaded: list[str], base_dir: str) -> dict:
    """
    Goi invoice_classifier.py de phan loai cac file PDF vua tai.
    Tra ve dict: {"classified": int, "unclassified": int, "error": str|None}
    """
    classifier_script = os.path.join(os.path.dirname(__file__), "invoice_classifier.py")
    if not os.path.exists(classifier_script):
        return {"classified": 0, "unclassified": 0, "error": f"Khong tim thay: {classifier_script}"}

    unclassified_base = os.path.join(base_dir, "Chua phan loai")
    files_arg = " ".join(f'"{f}"' for f in downloaded)
    cmd = [
        sys.executable, classifier_script,
        "--files", *downloaded,
        "--unclassified-base", unclassified_base
    ]

    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, encoding="utf-8", timeout=120
        )
        if not result.stdout.strip():
            return {"classified": 0, "unclassified": 0, "error": result.stderr.strip() or "Script khong tra ve ket qua"}
        data = json.loads(result.stdout)
        if not data.get("success"):
            return {"classified": 0, "unclassified": 0, "error": data.get("error", "Loi phan loai")}
        summary = data.get("summary", {})
        return {
            "classified": summary.get("classified", 0),
            "unclassified": summary.get("unclassified", 0),
            "error": None
        }
    except subprocess.TimeoutExpired:
        return {"classified": 0, "unclassified": 0, "error": "Timeout sau 120 giay"}
    except Exception as e:
        return {"classified": 0, "unclassified": 0, "error": str(e)}


def process_msg(msg, base_dir: str, log_file: str,
                success_log_file: str, forward_to: list[str]) -> bool:
    """
    Tai PDF dinh kem, chuyen tiep email, danh dau da doc.
    Tra ve True neu co file duoc tai.
    """
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

        if not downloaded:
            info = {
                "event": "no_pdf",
                "time": datetime.now().isoformat(),
                "subject": subject,
                "sender": sender,
                "attachments_total": att_count
            }
            _safe_print(info)
            _append_log(log_file, info)
            return False

        # Chuyen tiep email
        fwd_ok, fwd_err = forward_email(msg, forward_to)

        # Danh dau da doc
        read_ok = mark_as_read(msg)

        # Phan loai hoa don sau khi tai
        classify_result = classify_downloaded_files(downloaded, base_dir)

        # Event ra stdout + monitor_log
        evt = {
            "event": "downloaded",
            "time": datetime.now().isoformat(),
            "subject": subject,
            "sender": sender,
            "files": downloaded,
            "forwarded": fwd_ok,
            "forward_error": fwd_err if not fwd_ok else None,
            "marked_read": read_ok,
            "classify_classified": classify_result["classified"],
            "classify_unclassified": classify_result["unclassified"],
            "classify_error": classify_result["error"]
        }
        _safe_print(evt)
        _append_log(log_file, evt)

        # Ghi success log vinh vien
        success_entry = {
            "time": datetime.now().isoformat(),
            "subject": subject,
            "sender": sender,
            "files": [os.path.basename(f) for f in downloaded],
            "file_paths": downloaded,
            "forwarded_to": forward_to if fwd_ok else [],
            "forwarded": fwd_ok,
            "forward_error": fwd_err if not fwd_ok else None,
            "marked_read": read_ok,
            "classify_classified": classify_result["classified"],
            "classify_unclassified": classify_result["unclassified"],
            "classify_error": classify_result["error"]
        }
        _append_success_log(success_log_file, success_entry)

        return True

    except Exception as e:
        err = {"event": "error", "time": datetime.now().isoformat(), "message": f"process_msg: {e}"}
        _safe_print(err)
        _append_log(log_file, err)
        return False


def poll_inbox(namespace, base_dir: str, log_file: str,
               success_log_file: str, forward_to: list[str],
               processed_ids: set):
    """Scan tat ca Inbox/Invoice folder, xu ly email chua xu ly."""
    new_count = 0
    folders = get_all_inbox_folders(namespace)
    for folder in folders:
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            # Invoice folder: quet tat ca; Inbox: gioi han INBOX_SCAN_LIMIT
            is_invoice = folder.Name.lower() == "invoice"
            scan_limit = items.Count if is_invoice else INBOX_SCAN_LIMIT
            for i in range(1, min(items.Count + 1, scan_limit + 1)):
                try:
                    msg = items[i]
                    if msg.Class != 43:
                        continue
                    entry_id = msg.EntryID
                    if entry_id in processed_ids:
                        continue
                    processed_ids.add(entry_id)
                    if process_msg(msg, base_dir, log_file, success_log_file, forward_to):
                        new_count += 1
                except (IndexError, KeyError):
                    # Email bi xoa/di chuyen trong luc duyet — binh thuong
                    break
                except Exception as e:
                    err = {"event": "error", "time": datetime.now().isoformat(), "message": f"poll item: {e}"}
                    _safe_print(err)
                    _append_log(log_file, err)
        except Exception as e:
            err = {"event": "error", "time": datetime.now().isoformat(), "message": f"poll_folder: {e}"}
            _safe_print(err)
            _append_log(log_file, err)
    return new_count


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--save-dir", required=True)
    parser.add_argument("--log-file", required=True)
    parser.add_argument("--success-log-file", default="")
    parser.add_argument("--forward-to", default="",
                        help="Danh sach email cach nhau dau phay, vd: ke.toan@cty.com,gd@cty.com")
    parser.add_argument("--pid-file", default="")
    args = parser.parse_args()

    os.makedirs(args.save_dir, exist_ok=True)

    # Default success log file neu khong truyen
    success_log_file = args.success_log_file or os.path.join(
        os.path.dirname(args.log_file), "invoice_success_log.json")

    forward_to = [e.strip() for e in args.forward_to.split(",") if e.strip()] if args.forward_to else []

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
        _safe_print({"event": "error", "message": f"Khong the ket noi Outlook: {e}"})
        sys.exit(1)

    _safe_print({
        "event": "started",
        "time": datetime.now().isoformat(),
        "pid": os.getpid(),
        "save_dir": args.save_dir,
        "forward_to": forward_to
    })

    _append_log(args.log_file, {
        "event": "started",
        "time": datetime.now().isoformat(),
        "pid": os.getpid(),
        "forward_to": forward_to
    })

    processed_ids: set = set()

    # Init scan: chi danh dau email DA DOC la da xu ly de tranh tai lai.
    # Email CHUA DOC (UnRead=True) se duoc xu ly o poll dau tien.
    try:
        for folder in get_all_inbox_folders(namespace):
            try:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)
                for i in range(1, items.Count + 1):
                    try:
                        msg = items[i]
                        if msg.Class == 43 and not msg.UnRead:
                            processed_ids.add(msg.EntryID)
                    except (IndexError, KeyError):
                        break
                    except Exception:
                        pass
            except Exception:
                pass
        info = {
            "event": "init",
            "time": datetime.now().isoformat(),
            "marked_read_as_processed": len(processed_ids)
        }
        _safe_print(info)
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
                _safe_print(poll_info)
                _append_log(args.log_file, poll_info)
                poll_inbox(namespace, args.save_dir, args.log_file,
                           success_log_file, forward_to, processed_ids)
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
        _safe_print({"event": "stopped", "time": datetime.now().isoformat()})


if __name__ == "__main__":
    main()
