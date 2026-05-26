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
            string dir = GetUserDataDir(settings);
            Directory.CreateDirectory(dir);

            var form = new Form
            {
                Text = "Đăng nhập Zalo Web — Đóng cửa sổ này sau khi đăng nhập xong",
                Size = new Size(960, 680),
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

            form.Show();
        }

        // ── Gửi một tin nhắn vào nhóm (tìm theo tên nhóm) ──────────────────
        // Phải gọi trên UI thread
        public static async Task<(bool ok, string error)> SendToGroupAsync(
            ZaloSettings settings, string groupName, string message)
        {
            string dir = GetUserDataDir(settings);
            Directory.CreateDirectory(dir);

            var tcs = new TaskCompletionSource<(bool, string)>();

            var form = new Form
            {
                Text          = $"Zalo — đang gửi thông báo đến '{groupName}'...",
                Size          = new Size(960, 680),
                StartPosition = FormStartPosition.Manual,
                Location      = new System.Drawing.Point(-9999, -9999), // ẩn ngoài màn hình
                ShowInTaskbar = false
            };
            var wv2 = new WebView2 { Dock = DockStyle.Fill };
            form.Controls.Add(wv2);

            form.FormClosed += (s, e) =>
                tcs.TrySetResult((false, "Cửa sổ bị đóng trước khi gửi xong."));

            form.Load += async (s, e) =>
            {
                try
                {
                    var env = await CoreWebView2Environment.CreateAsync(null, dir);
                    await wv2.EnsureCoreWebView2Async(env);

                    // Hook event sau khi CoreWebView2 sẵn sàng
                    bool handled = false;
                    wv2.NavigationCompleted += async (s2, e2) =>
                    {
                        if (handled) return;
                        handled = true;

                        await Task.Delay(5000); // Chờ Zalo SPA render xong

                        try
                        {
                            var result = await PerformSendAsync(wv2, groupName, message);
                            tcs.TrySetResult(result);

                            if (result.ok)
                            {
                                await Task.Delay(800);
                                form.BeginInvoke(() => form.Close());
                            }
                            else
                            {
                                form.BeginInvoke(() =>
                                    form.Text = $"❌ {result.error}  —  Đóng cửa sổ khi xong");
                            }
                        }
                        catch (Exception ex)
                        {
                            tcs.TrySetResult((false, ex.Message));
                            form.BeginInvoke(() => form.Close());
                        }
                    };

                    wv2.Source = new Uri("https://chat.zalo.me/");
                }
                catch (Exception ex)
                {
                    tcs.TrySetResult((false, $"Lỗi WebView2: {ex.Message}"));
                    form.Close();
                }
            };

            form.Show();
            return await tcs.Task;
        }

        // ── Tự động tìm nhóm và gửi tin nhắn bằng JavaScript ───────────────
        private static async Task<(bool ok, string error)> PerformSendAsync(
            WebView2 wv2, string groupName, string message)
        {
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

            // Gõ tên nhóm vào ô tìm kiếm
            string safe = EscapeJs(groupName);
            string searchRes = await wv2.ExecuteScriptAsync($@"
                (function() {{
                    var inputs = document.querySelectorAll('input');
                    for (var i = 0; i < inputs.length; i++) {{
                        var el = inputs[i];
                        if (el.offsetParent === null) continue;
                        el.focus();
                        var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                        setter.call(el, '{safe}');
                        el.dispatchEvent(new Event('input',  {{ bubbles: true }}));
                        el.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        return 'ok';
                    }}
                    return 'no_input';
                }})()");

            if (searchRes.Trim('"') == "no_input")
                return (false, "Không tìm thấy ô tìm kiếm trên Zalo Web.");

            await Task.Delay(2200); // Chờ kết quả tìm kiếm

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

            await Task.Delay(1500);

            // Gõ tin nhắn vào ô contenteditable
            string safeMsg = EscapeJs(message);
            string typeRes = await wv2.ExecuteScriptAsync($@"
                (function() {{
                    var els = document.querySelectorAll('div[contenteditable=""true""]');
                    for (var el of els) {{
                        if (el.offsetParent === null) continue;
                        el.focus();
                        document.execCommand('insertText', false, '{safeMsg}');
                        return 'ok';
                    }}
                    return 'no_input';
                }})()");

            if (typeRes.Trim('"') == "no_input")
                return (false, "Không tìm thấy ô nhập tin nhắn.");

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

            await Task.Delay(1200);
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

                // Delay giữa các NCC để tránh bị Zalo giới hạn / khóa tài khoản
                if (!first) await Task.Delay(8000);
                first = false;

                await SendToGroupAsync(settings, sup.Zalo_Group_ID, messageText);
            }
        }
    }
}
