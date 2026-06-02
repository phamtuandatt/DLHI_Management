"""
Invoice Link Downloader
- Quét email trong folder Outlook (không có file PDF đính kèm)
- Phát hiện link hóa đơn điện tử (ehoadon.vn, meinvoice.vn, minvoice.misa.vn)
- Trích xuất mã tra cứu từ nội dung email
- Tự động tải PDF về bằng Selenium headless Chrome
- Trả về JSON kết quả
"""

import os
import sys
import io
import re
import json
import time
import shutil
import argparse
import tempfile
from datetime import datetime
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

try:
    import win32com.client
    import pythoncom
except ImportError:
    print(json.dumps({"success": False, "error": "pywin32 chưa cài"}))
    sys.exit(1)

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError:
    print(json.dumps({"success": False, "error": "selenium/webdriver-manager chưa cài"}))
    sys.exit(1)

OUTLOOK_FOLDER_NAME = "Invoice"

# ── Pattern nhận dạng loại hóa đơn và mã tra cứu ─────────────────────────────

PROVIDERS = {
    "ehoadon": {
        "url_patterns": [r"tracuu\.ehoadon\.vn", r"tchd\.ehoadon\.vn", r"ehoadon\.vn"],
        "code_patterns": [
            # MTC= trong URL
            r"MTC=([A-Z0-9]{6,20})",
            # tracuu.ehoadon.vn/<code>
            r"tracuu\.ehoadon\.vn/([A-Z0-9]{6,20})",
            # "Mã tra cứu là: XXXXX"
            r"[Mm]ã\s+tra\s+c[uứ]u\s+l[aà][:\s*\*]+([A-Z0-9]{6,20})",
            r"[Mm]ã\s+tra\s+c[uứ]u[:\s]+([A-Z0-9]{6,20})",
            r"(LSKH[A-Z0-9]{6,16})",
        ],
        "name": "eHoadon (BKAV)"
    },
    "meinvoice": {
        "url_patterns": [r"meinvoice\.vn", r"minvoice\.misa\.vn"],
        "code_patterns": [
            r"[Nn]h[aậ]p\s+m[aã]\s+s[oố][:\s]+([A-Z0-9_\-]{6,40})",
            r"m[aã]\s+s[oố][:\s]+([A-Z0-9_\-]{6,40})",
            r"lookup\s*code[:\s]+([A-Z0-9_\-]{6,40})",
        ],
        "name": "MISA meInvoice"
    },
    "minvoice": {
        # Email từ no-reply@kiemtrahoadon.vn, link tracuuhoadon.minvoice.vn
        "url_patterns": [r"tracuuhoadon\.minvoice\.vn", r"kiemtrahoadon\.vn", r"minvoice\.vn"],
        # Cần trích xuất 2 trường: mã số thuế bên bán + số bảo mật
        "code_patterns": [
            # Số bảo mật / mã tra cứu (16+ ký tự hex)
            r"[Ss][oố]\s+b[aả]o\s+m[aậ]t[^:]*:\s*([A-F0-9]{10,32})",
            r"[Mm]ã\s+[Ss][oố]\s+b[aả]o\s+m[aậ]t[^:]*:\s*([A-F0-9]{10,32})",
            r"[Mm]ã\s+tra\s+c[uứ]u\s+h[oó]a\s+đ[oơ]n[^:]*:\s*([A-F0-9]{10,32})",
        ],
        "tax_patterns": [
            # Mã số thuế bên bán (đơn vị bán)
            r"[Mm]ã\s+s[oố]\s+thu[eế].*?[bá]n[^:]*:\s*(\d{10,13})",
            r"[Bb][ướ][cC]\s+2[^:]*:\s*(\d{10,13})",
            r"[Tt]hu[eế].*?[bá]n[^:\d]*(\d{10,13})",
        ],
        "name": "M-Invoice (minvoice.vn)"
    },
}


def detect_provider(text: str, html: str) -> dict | None:
    """Nhận dạng nhà cung cấp từ nội dung email."""
    combined = (text + " " + html).lower()
    for key, info in PROVIDERS.items():
        for pat in info["url_patterns"]:
            if re.search(pat, combined, re.IGNORECASE):
                return {**info, "provider_key": key}
    return None


def extract_lookup_code(text: str, html: str, provider: dict) -> str | None:
    """Trích xuất mã tra cứu từ nội dung email."""
    # Ưu tiên plain text, fallback sang HTML (strip tags)
    sources = [text, re.sub(r"<[^>]+>", " ", html)]
    for source in sources:
        for pat in provider["code_patterns"]:
            m = re.search(pat, source, re.IGNORECASE)
            if m:
                code = m.group(1) if m.lastindex else m.group(0)
                return code.strip()
    return None


def extract_seller_tax(text: str, html: str, provider: dict) -> str | None:
    """Trích xuất mã số thuế bên bán (dùng cho minvoice)."""
    patterns = provider.get("tax_patterns", [])
    sources = [text, re.sub(r"<[^>]+>", " ", html)]
    for source in sources:
        for pat in patterns:
            m = re.search(pat, source, re.IGNORECASE)
            if m:
                return m.group(1).strip()
    return None


def extract_direct_url(text: str, html: str, provider: dict) -> str | None:
    """Tìm URL tra cứu trực tiếp trong email."""
    combined = text + " " + html
    for pat in provider["url_patterns"]:
        # Tìm URL đầy đủ http(s)://...ehoadon.vn/...
        full_pat = r'https?://[^\s"<>]*' + pat.replace(r"\.", r"\.") + r'[^\s"<>]*'
        m = re.search(full_pat, combined, re.IGNORECASE)
        if m:
            return m.group(0).rstrip(".,;)")
    return None


# ── Selenium helpers ───────────────────────────────────────────────────────────

def make_driver(download_dir: str) -> webdriver.Chrome:
    """Chrome headless — dùng cho meinvoice.vn."""
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1280,900")
    options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    })
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def make_edge_driver(download_dir: str):
    """Edge visible (ẩn ngoài màn hình) — dùng cho minvoice.vn."""
    from selenium.webdriver.edge.service import Service as EdgeService
    from selenium.webdriver.edge.options import Options as EdgeOptions
    from webdriver_manager.microsoft import EdgeChromiumDriverManager
    opts = EdgeOptions()
    opts.add_argument("--window-position=-2000,0")
    opts.add_argument("--window-size=1280,900")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-notifications")
    opts.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.protocol_handlers": 1,
    })
    return webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=opts)


def wait_for_download(download_dir: str, timeout: int = 30) -> str | None:
    """Chờ file PDF xuất hiện trong thư mục download."""
    end = time.time() + timeout
    while time.time() < end:
        files = [f for f in os.listdir(download_dir)
                 if f.lower().endswith(".pdf") and not f.endswith(".crdownload")]
        if files:
            # Lấy file mới nhất
            return max(
                [os.path.join(download_dir, f) for f in files],
                key=os.path.getmtime
            )
        time.sleep(0.8)
    return None


def download_ehoadon(lookup_code: str, download_dir: str) -> str | None:
    """Tải PDF từ tracuu.ehoadon.vn — navigate trực tiếp bằng MTC code, lấy PDF URL từ network log."""
    import json as _json
    import requests as _req

    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--window-size=1280,900")
    opts.set_capability("goog:loggingPrefs", {"performance": "ALL"})
    opts.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
    })
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    driver.execute_cdp_cmd("Network.enable", {})

    try:
        driver.get(f"http://tracuu.ehoadon.vn/{lookup_code}")
        time.sleep(10)

        # Thu thập PDF URL từ network log
        pdf_urls = []
        for entry in driver.get_log("performance"):
            try:
                msg = _json.loads(entry["message"])["message"]
                if msg.get("method") == "Network.requestWillBeSent":
                    url = msg["params"]["request"]["url"]
                    if "pdf" in url.lower():
                        pdf_urls.append(url)
                elif msg.get("method") == "Network.responseReceived":
                    url = msg["params"]["response"]["url"]
                    mime = msg["params"]["response"].get("mimeType", "")
                    if "pdf" in mime.lower() or "pdf" in url.lower():
                        pdf_urls.append(url)
            except Exception:
                pass

        # Tìm thêm trong iframe frameViewInvoice
        try:
            driver.switch_to.frame(driver.find_element(By.ID, "frameViewInvoice"))
            time.sleep(3)
            for tag in ["embed", "object", "iframe"]:
                for el in driver.find_elements(By.TAG_NAME, tag):
                    src = el.get_attribute("src") or el.get_attribute("data") or ""
                    if "pdf" in src.lower():
                        pdf_urls.append(src)
            src_text = driver.page_source
            pdf_urls += re.findall(r'https?://[^\s"\'<>]+\.pdf[^\s"\'<>]*', src_text, re.IGNORECASE)
            driver.switch_to.default_content()
        except Exception:
            driver.switch_to.default_content()

        # Download PDF
        if pdf_urls:
            cookies = {c["name"]: c["value"] for c in driver.get_cookies()}
            for url in dict.fromkeys(pdf_urls):  # deduplicate, preserve order
                url = url.strip("\"'")
                r = _req.get(url, cookies=cookies, timeout=20,
                             headers={"Referer": "https://tchd.ehoadon.vn/"})
                if r.status_code == 200 and len(r.content) > 5000:
                    dest = os.path.join(download_dir, f"invoice_{lookup_code}.pdf")
                    with open(dest, "wb") as f:
                        f.write(r.content)
                    return dest

        return wait_for_download(download_dir, timeout=5)
    except Exception:
        return None
    finally:
        driver.quit()


def download_meinvoice(lookup_code: str, download_dir: str) -> str | None:
    """Tải PDF từ meinvoice.vn/tra-cuu bằng mã tra cứu."""
    from selenium.webdriver.common.keys import Keys
    import requests

    driver = make_driver(download_dir)
    try:
        driver.get("https://www.meinvoice.vn/tra-cuu")
        wait = WebDriverWait(driver, 15)

        # Nhập mã vào #txtCode và nhấn Enter (kính lúp)
        input_box = wait.until(EC.presence_of_element_located((By.ID, "txtCode")))
        input_box.clear()
        input_box.send_keys(lookup_code)
        input_box.send_keys(Keys.RETURN)

        # Chờ iframe frmResult load xong với URL DownloadHandler
        time.sleep(8)

        # Cách 1: Lấy URL download trực tiếp từ iframe frmResult src
        for frame in driver.find_elements(By.TAG_NAME, "iframe"):
            src = frame.get_attribute("src") or ""
            if "DownloadHandler" in src and "Code=" in src:
                cookies = {c["name"]: c["value"] for c in driver.get_cookies()}
                resp = requests.get(src, cookies=cookies, timeout=30,
                                    headers={"Referer": "https://www.meinvoice.vn/"})
                if resp.status_code == 200 and len(resp.content) > 1000:
                    dest = os.path.join(download_dir, f"invoice_{lookup_code}.pdf")
                    with open(dest, "wb") as f:
                        f.write(resp.content)
                    return dest

        # Cách 2: Click nút "Tải hóa đơn dạng PDF" → chờ file download
        pdf_btns = driver.find_elements(By.CSS_SELECTOR, ".dm-item.pdf")
        if pdf_btns:
            driver.execute_script("arguments[0].click();", pdf_btns[0])
            return wait_for_download(download_dir, timeout=20)

        return None
    except Exception:
        return None
    finally:
        driver.quit()


def download_minvoice(lookup_code: str, download_dir: str, seller_tax: str = "") -> str | None:
    """Tải PDF từ tracuuhoadon.minvoice.vn — dùng Edge, xử lý tab quảng cáo và permission dialog."""
    driver = make_edge_driver(download_dir)
    main_window = driver.current_window_handle

    def close_extra_tabs():
        for handle in list(driver.window_handles):
            if handle == main_window:
                continue
            try:
                driver.switch_to.window(handle)
                url = driver.current_url
                if "permission-request-dialog" in url:
                    try:
                        allow_btn = driver.find_element(By.XPATH,
                            "//button[contains(.,'Allow') or contains(.,'Cho phép')]"
                            " | //*[@id='allow']")
                        allow_btn.click()
                        time.sleep(1)
                    except Exception:
                        pass
                else:
                    driver.close()
            except Exception:
                pass
        driver.switch_to.window(main_window)

    try:
        driver.get("http://tracuuhoadon.minvoice.vn/")
        time.sleep(4)
        close_extra_tabs()

        # Lấy tất cả input visible
        inputs = [i for i in driver.find_elements(By.CSS_SELECTOR, "input[type='text'], input:not([type])")
                  if i.is_displayed() and i.is_enabled()]

        if len(inputs) >= 1 and seller_tax:
            inputs[0].clear()
            inputs[0].send_keys(seller_tax)
        if len(inputs) >= 2:
            inputs[1].clear()
            inputs[1].send_keys(lookup_code)
        elif len(inputs) == 1:
            inputs[0].clear()
            inputs[0].send_keys(lookup_code)

        # Click Tra cứu, thử lại tối đa 3 lần nếu gặp tab quảng cáo
        btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(.,'Tra') or contains(.,'cứu')]"
                       " | //input[@value='Tra cứu']")))

        for attempt in range(3):
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(4)
            close_extra_tabs()
            # Kiểm tra có kết quả chưa
            pdf_btns = driver.find_elements(By.XPATH,
                "//*[contains(text(),'PDF') or contains(@class,'pdf')]")
            if pdf_btns:
                break

        # Click nút PDF
        pdf_btns = driver.find_elements(By.XPATH,
            "//*[contains(text(),'PDF') or contains(@class,'pdf')]")
        if pdf_btns:
            driver.execute_script("arguments[0].click();", pdf_btns[0])
            time.sleep(3)
            close_extra_tabs()

        return wait_for_download(download_dir, timeout=30)
    except Exception:
        return None
    finally:
        try:
            driver.quit()
        except Exception:
            pass


DOWNLOADERS = {
    "ehoadon": download_ehoadon,
    "meinvoice": download_meinvoice,
    "minvoice": download_minvoice,
}


# ── Xử lý từng email ──────────────────────────────────────────────────────────

def has_pdf_attachment(msg) -> bool:
    for att in msg.Attachments:
        if att.FileName.lower().endswith(".pdf"):
            return True
    return False


def process_email(msg, save_dir: str) -> dict:
    subject = msg.Subject or ""
    sender = msg.SenderName or msg.SenderEmailAddress or "Unknown"

    try:
        body_text = msg.Body or ""
        body_html = msg.HTMLBody or ""
    except Exception:
        body_text = ""
        body_html = ""

    provider = detect_provider(body_text, body_html)
    if not provider:
        return {
            "status": "no_link",
            "subject": subject,
            "sender": sender,
            "reason": "Không tìm thấy link hóa đơn điện tử trong email"
        }

    lookup_code = extract_lookup_code(body_text, body_html, provider)
    if not lookup_code:
        seller_tax_dbg = extract_seller_tax(body_text, body_html, provider) or ""
        return {
            "status": "pending_manual",
            "subject": subject,
            "sender": sender,
            "provider": provider["name"],
            "seller_tax": seller_tax_dbg,
            "reason": "Không trích xuất được mã tra cứu — cần xử lý thủ công"
        }

    # Trích xuất mã số thuế bên bán nếu cần (minvoice)
    seller_tax = extract_seller_tax(body_text, body_html, provider) or ""

    # Tải PDF vào thư mục tạm rồi move sang save_dir
    with tempfile.TemporaryDirectory() as tmp_dir:
        downloader = DOWNLOADERS.get(provider["provider_key"])
        if provider["provider_key"] == "minvoice":
            downloaded = downloader(lookup_code, tmp_dir, seller_tax) if downloader else None
        else:
            downloaded = downloader(lookup_code, tmp_dir) if downloader else None

        if not downloaded:
            return {
                "status": "pending_manual",
                "subject": subject,
                "sender": sender,
                "provider": provider["name"],
                "lookup_code": lookup_code,
                "reason": "Tải thất bại (captcha hoặc lỗi trang) — cần xử lý thủ công"
            }

        # Đặt tên file: INV_<lookup_code>.pdf
        safe_code = re.sub(r'[\\/:*?"<>|]', "_", lookup_code)
        dest_name = f"INV_{safe_code}.pdf"
        dest_path = os.path.join(save_dir, dest_name)
        if os.path.exists(dest_path):
            dest_path = os.path.join(save_dir, f"INV_{safe_code}_{int(time.time())}.pdf")

        shutil.move(downloaded, dest_path)

    return {
        "status": "downloaded",
        "subject": subject,
        "sender": sender,
        "provider": provider["name"],
        "lookup_code": lookup_code,
        "file": dest_path
    }


def scan_emails(save_dir: str, days_back: int) -> list[dict]:
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
    except Exception as e:
        return [{"status": "error", "reason": f"Không kết nối Outlook: {e}"}]

    # Tìm folder Invoice
    folder = None
    for store in ns.Stores:
        try:
            root = store.GetRootFolder()
            for f in root.Folders:
                if f.Name.lower() == OUTLOOK_FOLDER_NAME.lower():
                    folder = f
                    break
                for sub in f.Folders:
                    if sub.Name.lower() == OUTLOOK_FOLDER_NAME.lower():
                        folder = sub
                        break
            if folder:
                break
        except Exception:
            pass

    if not folder:
        return [{"status": "error", "reason": f"Không tìm thấy folder '{OUTLOOK_FOLDER_NAME}'"}]

    os.makedirs(save_dir, exist_ok=True)
    cutoff = datetime.now().timestamp() - days_back * 86400
    results = []

    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)

    for msg in messages:
        try:
            if msg.Class != 43:
                continue
            if msg.ReceivedTime.timestamp() < cutoff:
                break
            if has_pdf_attachment(msg):
                continue  # Đã có PDF đính kèm, bỏ qua
            result = process_email(msg, save_dir)
            if result["status"] != "no_link":
                results.append(result)
        except Exception as e:
            results.append({"status": "error", "reason": str(e)})

    return results


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--save-dir", required=True)
    parser.add_argument("--days-back", type=int, default=30)
    args = parser.parse_args()

    results = scan_emails(args.save_dir, args.days_back)

    downloaded = sum(1 for r in results if r["status"] == "downloaded")
    pending = sum(1 for r in results if r["status"] == "pending_manual")

    print(json.dumps({
        "success": True,
        "results": results,
        "summary": {
            "total": len(results),
            "downloaded": downloaded,
            "pending_manual": pending
        }
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
