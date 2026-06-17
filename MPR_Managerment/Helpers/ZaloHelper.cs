using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using MPR_Managerment.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MPR_Managerment.Helpers
{
    public class ZaloSettings
    {
        public bool   Enabled     { get; set; } = false;
        public string UserDataDir { get; set; } = "";
    }

    public enum ZaloStatus { Disabled, Connecting, Ready, Error }

    // ── Phiên WebView2 duy trì suốt vòng đời ứng dụng ──────────────────────
    public static class ZaloSession
    {
        private static Form?    _form;
        private static WebView2? _wv2;
        private static System.Windows.Forms.Timer? _keepAlive;
        private static bool _ready = false;
        private static ZaloStatus _lastFired = ZaloStatus.Disabled;

        public static bool      IsReady => _ready && _wv2 != null && !_wv2.IsDisposed;
        public static WebView2? WebView => _wv2;

        // Đăng ký để nhận thông báo khi trạng thái Zalo thay đổi
        public static event Action<ZaloStatus>? StatusChanged;

        private static void _Fire(ZaloStatus s)
        {
            if (_lastFired == s) return;
            _lastFired = s;
            StatusChanged?.Invoke(s);
        }

        // Gọi trên UI thread — khởi tạo nếu chưa có session
        public static async Task<(bool ok, string error)> EnsureAsync(ZaloSettings settings)
        {
            if (IsReady) { _Fire(ZaloStatus.Ready); return (true, ""); }

            _Fire(ZaloStatus.Connecting);

            string dir = string.IsNullOrWhiteSpace(settings.UserDataDir)
                ? ZaloHelper.DefaultUserDataDir : settings.UserDataDir;
            Directory.CreateDirectory(dir);

            _form = new Form
            {
                Text          = "Zalo Background Session",
                Size          = new System.Drawing.Size(960, 680),
                StartPosition = FormStartPosition.Manual,
                Location      = new System.Drawing.Point(-9999, -9999),
                ShowInTaskbar = false
            };
            _wv2 = new WebView2 { Dock = DockStyle.Fill };
            _form.Controls.Add(_wv2);

            _form.FormClosed += (s, e) =>
            {
                _ready = false;
                _keepAlive?.Stop();
                _keepAlive?.Dispose();
                _keepAlive = null;
                _wv2  = null;
                _form = null;
                _Fire(ZaloStatus.Error);
            };

            _form.Show();

            try
            {
                var env = await CoreWebView2Environment.CreateAsync(null, dir);
                await _wv2.EnsureCoreWebView2Async(env);

                var navTcs = new TaskCompletionSource<bool>();
                void Handler(object? s, CoreWebView2NavigationCompletedEventArgs e)
                {
                    _wv2!.NavigationCompleted -= Handler;
                    navTcs.TrySetResult(e.IsSuccess);
                }
                _wv2.NavigationCompleted += Handler;
                _wv2.Source = new Uri("https://chat.zalo.me/");

                await Task.WhenAny(navTcs.Task, Task.Delay(20_000));
                await Task.Delay(4_500); // chờ SPA render xong

                _StartKeepAlive();
                _ready = true;
                _Fire(ZaloStatus.Ready);
                return (true, "");
            }
            catch (Exception ex)
            {
                _ready = false;
                _Fire(ZaloStatus.Error);
                return (false, $"Lỗi WebView2: {ex.Message}");
            }
        }

        // Tự động bấm "Kích hoạt" nếu Zalo hiện dialog "đang mở trên tab khác"
        public static async Task<bool> TryDismissActivationDialogAsync(WebView2 wv2)
        {
            try
            {
                string res = await wv2.ExecuteScriptAsync(@"
                    (function() {
                        var btns = document.querySelectorAll('button');
                        for (var i = 0; i < btns.length; i++) {
                            var t = (btns[i].innerText || btns[i].textContent || '').trim();
                            if (t.includes('ch ho') || t.includes('Activate')) {
                                btns[i].click();
                                return 'clicked';
                            }
                        }
                        return 'none';
                    })()");
                return res.Trim('"') == "clicked";
            }
            catch { return false; }
        }

        // Keep-alive kiểm tra login và tự kích hoạt nếu bị dialog chặn — mỗi 2 phút
        private static void _StartKeepAlive()
        {
            _keepAlive = new System.Windows.Forms.Timer { Interval = 120_000 };
            _keepAlive.Tick += async (s, e) =>
            {
                try
                {
                    if (_wv2 == null || _wv2.IsDisposed) return;

                    // Tự kích hoạt nếu đang bị dialog "tab khác" chặn
                    await TryDismissActivationDialogAsync(_wv2);
                    await Task.Delay(1500);

                    string result = await _wv2.ExecuteScriptAsync(@"
                        (function() {
                            var inputs = document.querySelectorAll('input');
                            for (var i = 0; i < inputs.length; i++)
                                if (inputs[i].offsetParent !== null) return 'ok';
                            return 'logged_out';
                        })()");

                    bool loggedIn = result.Trim('"') == "ok";
                    _ready = loggedIn;
                    _Fire(loggedIn ? ZaloStatus.Ready : ZaloStatus.Error);
                }
                catch
                {
                    _ready = false;
                    _Fire(ZaloStatus.Error);
                }
            };
            _keepAlive.Start();
        }

        // Gọi trước khi mở login browser (tránh xung đột userDataDir)
        public static void Shutdown()
        {
            _ready = false;
            _keepAlive?.Stop();
            _keepAlive?.Dispose();
            _keepAlive = null;
            try { _form?.Close(); } catch { }
            _Fire(ZaloStatus.Disabled);
        }
    }

    public static class ZaloHelper
    {
        private const string SETTINGS_FILE = "zalo_settings.json";

        private static readonly string SettingsPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, SETTINGS_FILE);

        private static readonly JsonSerializerOptions _json = new() { WriteIndented = true };

        public static string DefaultUserDataDir => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ERP_ZaloBrowser");

        // ── Settings ─────────────────────────────────────────────────────────
        public static ZaloSettings LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsPath))
                    return JsonSerializer.Deserialize<ZaloSettings>(File.ReadAllText(SettingsPath))
                           ?? new ZaloSettings();
            }
            catch { }
            return new ZaloSettings();
        }

        public static void SaveSettings(ZaloSettings s)
            => File.WriteAllText(SettingsPath, JsonSerializer.Serialize(s, _json));

        private static string GetUserDataDir(ZaloSettings s)
            => string.IsNullOrWhiteSpace(s.UserDataDir) ? DefaultUserDataDir : s.UserDataDir;

        // ── Mở trình duyệt để đăng nhập lần đầu ────────────────────────────
        public static void OpenBrowserForLogin(ZaloSettings settings)
        {
            // Phải tắt session nền trước vì cả hai không thể dùng cùng userDataDir
            ZaloSession.Shutdown();

            string dir = GetUserDataDir(settings);
            Directory.CreateDirectory(dir);

            var form = new Form
            {
                Text = "Đăng nhập Zalo Web — Đóng cửa sổ này sau khi đăng nhập xong",
                Size = new System.Drawing.Size(960, 680),
                StartPosition = FormStartPosition.CenterScreen
            };
            var wv2 = new WebView2 { Dock = DockStyle.Fill };
            form.Controls.Add(wv2);

            form.Load += async (s, e) =>
            {
                try
                {
                    var env = await CoreWebView2Environment.CreateAsync(null, dir);
                    await wv2.EnsureCoreWebView2Async(env);
                    wv2.Source = new Uri("https://chat.zalo.me/");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(form, $"Lỗi WebView2: {ex.Message}", "Lỗi");
                    form.Close();
                }
            };

            // Khi đóng login browser: session nền sẽ được khởi tạo lại vào lần gửi kế tiếp
            // (ZaloSession.IsReady = false sau Shutdown)
            form.Show();
        }

        // ── Gửi một tin nhắn vào nhóm (tìm theo tên nhóm) ──────────────────
        // Phải gọi trên UI thread
        public static async Task<(bool ok, string error)> SendToGroupAsync(
            ZaloSettings settings, string groupName, string message)
        {
            // ── Kiểm tra quyền gửi tin nhắn Zalo ──
            if (!AppSession.IsAdmin && !AppSession.HasPermission("ZALO", "Gửi tin nhắn"))
                return (false, "Bạn không có quyền gửi tin nhắn Zalo. Vui lòng liên hệ Admin để được cấp quyền.");

            // Đảm bảo session nền đã chạy
            if (!ZaloSession.IsReady)
            {
                var (initOk, initErr) = await ZaloSession.EnsureAsync(settings);
                if (!initOk) return (false, initErr);
            }

            var wv2 = ZaloSession.WebView;
            if (wv2 == null) return (false, "WebView2 không khả dụng");

            try
            {
                return await PerformSendAsync(wv2, groupName, message);
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        // ── Tự động tìm nhóm và gửi tin nhắn bằng JavaScript ───────────────
        private static async Task<(bool ok, string error)> PerformSendAsync(
            WebView2 wv2, string groupName, string message)
        {
            // Tự kích hoạt nếu Zalo đang hiện dialog "tab khác" trước khi làm bất cứ điều gì
            bool dismissed = await ZaloSession.TryDismissActivationDialogAsync(wv2);
            if (dismissed) await Task.Delay(2500); // chờ Zalo reload lại sau kích hoạt

            // Kiểm tra đã đăng nhập chưa
            string loginCheck = await wv2.ExecuteScriptAsync(@"
                (function() {
                    var inputs = document.querySelectorAll('input');
                    for (var i = 0; i < inputs.length; i++)
                        if (inputs[i].offsetParent !== null) return 'ok';
                    return 'not_logged_in';
                })()");

            if (loginCheck.Trim('"') == "not_logged_in")
                return (false, "Chưa đăng nhập Zalo Web. Hãy bấm 'Mở trình duyệt & đăng nhập' trước.");

            // Xoá tìm kiếm cũ, gõ tên nhóm mới vào ô tìm kiếm
            string safe = EscapeJs(groupName);
            string searchRes = await wv2.ExecuteScriptAsync($@"
                (function() {{
                    var inputs = document.querySelectorAll('input');
                    for (var i = 0; i < inputs.length; i++) {{
                        var el = inputs[i];
                        if (el.offsetParent === null) continue;
                        el.focus();
                        var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                        setter.call(el, '');
                        el.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        setter.call(el, '{safe}');
                        el.dispatchEvent(new Event('input',  {{ bubbles: true }}));
                        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        return 'ok';
                    }}
                    return 'no_input';
                }})()");

            if (searchRes.Trim('"') == "no_input")
                return (false, "Không tìm thấy ô tìm kiếm trên Zalo Web.");

            // Poll cho đến khi kết quả tìm kiếm xuất hiện (tối đa 6s, poll mỗi 300ms)
            // delay cứng 2200ms thiếu khi Zalo web đang re-render danh sách
            bool searchReady = false;
            for (int w = 0; w < 20 && !searchReady; w++)
            {
                await Task.Delay(300);
                string probe = await wv2.ExecuteScriptAsync(@"
                    (function() {
                        var ss = [
                            'div[class*=""conv-item""]',
                            'div[class*=""conversation-item""]',
                            'div[class*=""item-chat""]',
                            'div[class*=""contact-item""]',
                            'div[class*=""listItem""]',
                            'div[class*=""chat-item""]',
                            '[class*=""result-item""]'
                        ];
                        for (var sel of ss) {
                            var items = document.querySelectorAll(sel);
                            for (var item of items) {
                                if (item.offsetParent !== null) return 'found';
                            }
                        }
                        return 'waiting';
                    })()");
                searchReady = probe.Trim('"') == "found";
            }

            // Click vào kết quả đầu tiên
            string clickRes = await wv2.ExecuteScriptAsync(@"
                (function() {
                    var ss = [
                        'div[class*=""conv-item""]',
                        'div[class*=""conversation-item""]',
                        'div[class*=""item-chat""]',
                        'div[class*=""contact-item""]',
                        'div[class*=""listItem""]',
                        'div[class*=""chat-item""]',
                        '[class*=""result-item""]'
                    ];
                    for (var sel of ss) {
                        var items = document.querySelectorAll(sel);
                        for (var item of items) {
                            if (item.offsetParent !== null) { item.click(); return 'ok'; }
                        }
                    }
                    return 'not_found';
                })()");

            if (clickRes.Trim('"') == "not_found")
                return (false, $"Không tìm thấy nhóm '{groupName}' trong Zalo. Kiểm tra lại tên nhóm.");

            // Đóng search overlay sau khi click — quan trọng khi gửi liên tiếp cùng một nhóm.
            // Nếu nhóm đó đang active, Zalo web không tự đóng overlay sau khi click,
            // khiến contenteditable bị overlay che và bị mất focus trap → Enter không gửi được.
            await wv2.ExecuteScriptAsync(@"
                (function() {
                    var inputs = document.querySelectorAll('input');
                    for (var i = 0; i < inputs.length; i++) {
                        var el = inputs[i];
                        if (el.offsetParent === null) continue;
                        var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                        setter.call(el, '');
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.blur();
                        return;
                    }
                })()");
            await Task.Delay(800);

            // Poll cho đến khi ô contenteditable xuất hiện (tối đa 8s, poll mỗi 300ms)
            // Delay cứng 1500ms không đủ khi Zalo web đang render lại chat panel
            string typeRes = "no_input";
            for (int waitStep = 0; waitStep < 27 && typeRes.Trim('"') != "ok"; waitStep++)
            {
                await Task.Delay(300);
                typeRes = await wv2.ExecuteScriptAsync(@"
                    (function() {
                        var els = document.querySelectorAll('div[contenteditable=""true""]');
                        for (var el of els) {
                            if (el.offsetParent === null) continue;
                            el.focus();
                            document.execCommand('selectAll', false, null);
                            document.execCommand('delete', false, null);
                            return 'ok';
                        }
                        return 'no_input';
                    })()");
            }

            if (typeRes.Trim('"') == "no_input")
            {
                // Zalo có thể đã đóng conversation panel sau khi clear search overlay.
                // Thử click lại vào conversation item trong sidebar trái (không qua search).
                string retryClick = await wv2.ExecuteScriptAsync(@"
                    (function() {
                        var ss = [
                            'div[class*=""conv-item""]',
                            'div[class*=""conversation-item""]',
                            'div[class*=""item-chat""]',
                            'div[class*=""contact-item""]',
                            'div[class*=""listItem""]',
                            'div[class*=""chat-item""]'
                        ];
                        for (var sel of ss) {
                            var items = document.querySelectorAll(sel);
                            for (var item of items) {
                                if (item.offsetParent !== null && item.getAttribute('data-active') !== null) {
                                    item.click(); return 'retry_clicked';
                                }
                            }
                            // Nếu không có data-active, click cái đầu tiên visible
                            for (var item of items) {
                                if (item.offsetParent !== null) { item.click(); return 'retry_fallback'; }
                            }
                        }
                        return 'none';
                    })()");
                await Task.Delay(1000);

                // Poll lần 2 sau khi retry click
                for (int waitStep2 = 0; waitStep2 < 15 && typeRes.Trim('"') != "ok"; waitStep2++)
                {
                    await Task.Delay(300);
                    typeRes = await wv2.ExecuteScriptAsync(@"
                        (function() {
                            var els = document.querySelectorAll('div[contenteditable=""true""]');
                            for (var el of els) {
                                if (el.offsetParent === null) continue;
                                el.focus();
                                document.execCommand('selectAll', false, null);
                                document.execCommand('delete', false, null);
                                return 'ok';
                            }
                            return 'no_input';
                        })()");
                }
            }

            if (typeRes.Trim('"') == "no_input")
                return (false, "Không tìm thấy ô nhập tin nhắn.");

            // Gửi từng dòng bằng Input.insertText, giữa các dòng dùng Shift+Enter để xuống dòng.
            // Gửi toàn bộ text (có \n) một lần khiến Zalo Web kích hoạt handler gửi tin tại \n đầu tiên
            // → chỉ dòng tiêu đề được gửi, phần còn lại mất. Shift+Enter (modifiers=8) tạo
            // xuống dòng trong ô nhập mà không trigger gửi tin.
            const string shiftEnterDown = @"{""type"":""keyDown"",""key"":""Enter"",""code"":""Enter"",""windowsVirtualKeyCode"":13,""nativeVirtualKeyCode"":13,""modifiers"":8}";
            const string shiftEnterChar = @"{""type"":""char"",""key"":""\r"",""code"":""Enter"",""windowsVirtualKeyCode"":13,""nativeVirtualKeyCode"":13,""modifiers"":8}";
            const string shiftEnterUp   = @"{""type"":""keyUp"",""key"":""Enter"",""code"":""Enter"",""windowsVirtualKeyCode"":13,""nativeVirtualKeyCode"":13,""modifiers"":8}";

            // Dùng UnsafeRelaxedJsonEscaping để emoji được giữ nguyên dạng literal UTF-8.
            // JsonSerializer mặc định escape emoji thành surrogate pairs (📅) —
            // CDP Input.insertText xử lý sai surrogate pairs, khiến emoji bị tách khỏi text.
            var cdpJsonOpts = new System.Text.Json.JsonSerializerOptions
            {
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            var lines = message.Split('\n');
            for (int li = 0; li < lines.Length; li++)
            {
                string lineText = lines[li].TrimEnd('\r');
                if (!string.IsNullOrEmpty(lineText))
                {
                    string cdpLineJson = System.Text.Json.JsonSerializer.Serialize(lineText, cdpJsonOpts);
                    await wv2.CoreWebView2.CallDevToolsProtocolMethodAsync(
                        "Input.insertText", $@"{{""text"":{cdpLineJson}}}");
                    await Task.Delay(50);
                }
                if (li < lines.Length - 1)
                {
                    await wv2.CoreWebView2.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", shiftEnterDown);
                    await Task.Delay(30);
                    await wv2.CoreWebView2.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", shiftEnterChar);
                    await Task.Delay(30);
                    await wv2.CoreWebView2.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", shiftEnterUp);
                    await Task.Delay(50);
                }
            }

            await Task.Delay(400);

            // Focus lại ô nhập trước khi gửi
            await wv2.ExecuteScriptAsync(@"
                (function() {
                    var els = document.querySelectorAll('div[contenteditable=""true""]');
                    for (var el of els) {
                        if (el.offsetParent !== null) { el.focus(); return; }
                    }
                })()");
            await Task.Delay(200);

            // Dùng CDP Input.dispatchKeyEvent — React nhận được, JS dispatchEvent thông thường không đủ
            const string cdpKeyDown = @"{""type"":""keyDown"",""key"":""Enter"",""code"":""Enter"",""windowsVirtualKeyCode"":13,""nativeVirtualKeyCode"":13,""modifiers"":0}";
            const string cdpChar    = @"{""type"":""char"",  ""key"":""\r"",    ""code"":""Enter"",""windowsVirtualKeyCode"":13,""nativeVirtualKeyCode"":13,""modifiers"":0}";
            const string cdpKeyUp   = @"{""type"":""keyUp"", ""key"":""Enter"",""code"":""Enter"",""windowsVirtualKeyCode"":13,""nativeVirtualKeyCode"":13,""modifiers"":0}";

            await wv2.CoreWebView2.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", cdpKeyDown);
            await Task.Delay(60);
            await wv2.CoreWebView2.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", cdpChar);
            await Task.Delay(60);
            await wv2.CoreWebView2.CallDevToolsProtocolMethodAsync("Input.dispatchKeyEvent", cdpKeyUp);

            // Chờ Zalo xử lý xong gửi tin và reset về trạng thái sẵn sàng
            // (2s đủ để message submit, tránh race condition với tin nhắn kế tiếp)
            await Task.Delay(2000);
            return (true, "");
        }

        private static string EscapeJs(string s)
            => (s ?? "")
               .Replace("\\", "\\\\")
               .Replace("'",  "\\'")
               .Replace("\"", "\\\"")
               .Replace("\r\n", "\\n")
               .Replace("\r",   "\\n")
               .Replace("\n",   "\\n");

        // ── Kiểm tra Zalo đã được cấu hình chưa ─────────────────────────────
        public static bool IsConfigured()
        {
            var s = LoadSettings();
            return s.Enabled && ZaloSession.IsReady;
        }

        // ── Gửi tin nhắn nhanh (dùng cho Dashboard) ─────────────────────────
        public static async void SendMessage(string message)
        {
            var settings = LoadSettings();
            if (!settings.Enabled) return;

            // Gửi vào nhóm mặc định "Giao hàng" nếu có
            string groupName = settings.UserDataDir; // fallback
            // Tìm nhóm "Giao hàng" hoặc nhóm đầu tiên có sẵn
            var (ok, err) = await SendToGroupAsync(settings, "Giao hàng", message);
            if (!ok)
            {
                System.Diagnostics.Debug.WriteLine($"ZaloHelper.SendMessage failed: {err}");
            }
        }

        // ── Gửi đến nhiều nhà cung cấp ──────────────────────────────────────
        public static async Task SendNotificationsToSuppliersAsync(
            IEnumerable<Supplier> suppliers, string messageText)
        {
            var settings = LoadSettings();
            if (!settings.Enabled) return;

            bool first = true;
            foreach (var sup in suppliers)
            {
                if (string.IsNullOrWhiteSpace(sup.Zalo_Group_ID)) continue;

                if (!first) await Task.Delay(8000);
                first = false;

                await SendToGroupAsync(settings, sup.Zalo_Group_ID, messageText);
            }
        }
    }
}
