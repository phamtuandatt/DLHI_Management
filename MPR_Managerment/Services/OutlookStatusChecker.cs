using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;

namespace MPR_Managerment.Services
{
    public record OutlookStatus(string Text, Color Color);

    public static class OutlookStatusChecker
    {
        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        private static object? GetActiveOutlook()
        {
            var type = Type.GetTypeFromProgID("Outlook.Application");
            if (type == null) return null;
            int hr = GetActiveObject(type.GUID, IntPtr.Zero, out object obj);
            return hr == 0 ? obj : null;
        }

        public static OutlookStatus Check()
        {
            // Bước 1: kiểm tra process Outlook có chạy không
            var procs = Process.GetProcessesByName("OUTLOOK");
            if (procs.Length == 0)
                return new OutlookStatus("Outlook: Chưa mở", Color.FromArgb(255, 80, 80));

            // Bước 2: thử kết nối COM để xác nhận MAPI sẵn sàng
            object? outlook = null;
            try
            {
                outlook = GetActiveOutlook();
                if (outlook == null)
                    return new OutlookStatus("Outlook: Đang khởi động...", Color.FromArgb(255, 200, 0));

                // Thử lấy namespace MAPI
                var ns = outlook.GetType().InvokeMember(
                    "GetNamespace", System.Reflection.BindingFlags.InvokeMethod,
                    null, outlook, new object[] { "MAPI" });

                if (ns == null)
                    return new OutlookStatus("Outlook: Chưa đăng nhập", Color.FromArgb(255, 160, 0));

                // Thử lấy CurrentUser để xác nhận tài khoản đã kết nối
                var user = ns.GetType().InvokeMember(
                    "CurrentUser", System.Reflection.BindingFlags.GetProperty,
                    null, ns, null);

                string? name = user?.GetType().InvokeMember(
                    "Name", System.Reflection.BindingFlags.GetProperty,
                    null, user, null) as string;

                return string.IsNullOrEmpty(name)
                    ? new OutlookStatus("Outlook: Đã kết nối", Color.FromArgb(0, 220, 110))
                    : new OutlookStatus($"Outlook: {name} ✓", Color.FromArgb(0, 220, 110));
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x800401E3))
            {
                // MK_E_UNAVAILABLE — Outlook đang mở nhưng chưa sẵn sàng
                return new OutlookStatus("Outlook: Đang khởi động...", Color.FromArgb(255, 200, 0));
            }
            catch
            {
                return procs.Length > 0
                    ? new OutlookStatus("Outlook: Đang chạy", Color.FromArgb(0, 180, 100))
                    : new OutlookStatus("Outlook: Lỗi kết nối", Color.FromArgb(255, 80, 80));
            }
            finally
            {
                if (outlook != null) Marshal.ReleaseComObject(outlook);
            }
        }
    }
}
