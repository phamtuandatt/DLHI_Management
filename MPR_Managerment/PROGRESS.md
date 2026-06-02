# PROGRESS — MPR_Managerment ERP

> Cập nhật: 2026-06-01  
> Dự án: `D:\ERP_Final_Git\MPR_Managerment\`  
> Stack: C# WinForms .NET 8 · SQL Server Azure · EPPlus (OfficeOpenXml)

---

## ✅ Task A — frmPayment UI Changes (HOÀN THÀNH)

**Yêu cầu gốc:**
- Bỏ bảng "History Paid" khỏi giao diện chính → chuyển thành popup khi click button
- Kéo dài panel "Payment Request Progressing" lấp đầy diện tích còn lại
- Xóa button + logic: "Thêm từ InRequest", "Lưu trạng thái", "Lưu thanh toán"

**Các thay đổi trong `Forms/frmPayment.cs`:**

| Thay đổi | Chi tiết |
|----------|----------|
| Xóa `panelPaid` field | Giữ `dgvPaid`, `_paidFrom`, `_paidTo` là standalone fields |
| Thêm button "📋 History Paid" | Trong toolbar panelHist, x=144, ngoài block `canEdit` |
| Giữ chỉ button "Xóa" | Đã di chuyển vào vị trí x=280 |
| `LoadData()` | Bỏ gọi `LoadHistoryPaid(...)` |
| `ResizeAll()` | `panelHist.Height = Math.Max(200, h - panelHist.Top - 10)` |
| Xóa 3 methods | `BtnSavePaymentStatus_Click`, `BtnSaveHistoryPaid_Click`, `BtnAddPayment_Click` |
| Thêm `ShowHistoryPaidPopup()` | Form popup 950×540, dgvPaid + bộ lọc dtpFrom/dtpTo |

**Popup History Paid features:**
- Filter theo khoảng ngày (dtpFrom → dtpTo)
- Button "Tất cả" reset về 2000-01-01 → hôm nay
- Button "Xóa" gọi `BtnDelHistoryPaid_Click`
- `dgvPaid.Parent = null` khi đóng popup để tái sử dụng

---

## ✅ Task B — Zalo Import Feature (HOÀN THÀNH — chờ confirm runtime)

**Yêu cầu gốc:**
- Auto-detect file `0. 품의서 YYYY-MM-DD.xlsx` từ folder Zalo Temp
- Đọc sheet "Payment list", từ row 4
- Import vào SQL Server Azure, ghi đè nếu trùng

### File đã tạo/sửa:

#### `Services/ZaloImportService.cs` (FILE MỚI)
- `ZaloDownloadFolder`: `%LOCALAPPDATA%\Temp\Zalo Temp\TempDownloads`
- `FindImportFiles()`: scan `0. 품의서 *.xlsx`, sort mới nhất trước, dùng `.Item1`/`.Item2`
- `ReadFile()`: đọc EPPlus, row 4+, skip dòng trống (kiểm tra col H + J + C + G)
- `ImportToDB()`: INSERT hoặc UPDATE theo `UNIQUE(File_Date, Row_Index)`
- `EnsureTable()`: tạo bảng `Zalo_PaymentImport` nếu chưa có

**Mapping cột Excel → DB:**

| DB Field | Excel Col | Index |
|----------|-----------|-------|
| PO_No | C | 3 |
| Ecount_No | D | 4 |
| Title_EN | E | 5 |
| GW_No | G | 7 |
| Amount | H | 8 |
| VAT | I | 9 |
| Final_Amount | J | 10 |
| Dot1–Dot8 | L–S | 12–19 |
| Progress_Status | AA | 27 |

**DB Schema (`Zalo_PaymentImport`):**
```sql
Import_ID INT IDENTITY PRIMARY KEY
File_Date DATE NOT NULL
Row_Index INT NOT NULL
PO_No, Ecount_No NVARCHAR
Title_EN NVARCHAR(500)
GW_No NVARCHAR(200)
Amount, VAT, Final_Amount DECIMAL(18,2)
Dot1–Dot8 DECIMAL(18,2)
Progress_Status NVARCHAR(100)
Imported_At DATETIME DEFAULT GETDATE()
Source_File NVARCHAR(500)
CONSTRAINT UQ_ZaloImport UNIQUE (File_Date, Row_Index)
```

#### `Forms/frmZaloImport.cs` (FILE MỚI)
- Code-only WinForms (KHÔNG dùng Designer, KHÔNG có `InitializeComponent()`)
- Layout: panelTop (Top, 82px) + lblWatch (Bottom) + panelStatus (Bottom) + SplitContainer (Fill)
- Thứ tự `Controls.Add`: Fill → Bottom → Bottom → Top (WinForms xử lý ngược)
- `SplitterDistance` set trong `Shown` event (KHÔNG phải `Load` hoặc constructor)
- Left panel: danh sách file (ListBox)
- Right panel: TabControl với 2 tab (Xem trước file / Dữ liệu DB)
- `FileSystemWatcher` + debounce timer 2000ms
- Auto-import khi phát hiện file mới (checkbox toggle)

#### `Forms/frmMain.cs` (SỬA)
- Thêm menu button: `📥  Zalo Import` → mở `frmZaloImport`

---

## 🐛 Lỗi đã gặp và đã sửa

| # | Lỗi | Nguyên nhân | Fix |
|---|-----|-------------|-----|
| 1 | `InitializeComponent` not found | frmZaloImport là code-only | Xóa lời gọi |
| 2 | `(string, DateTime)` không có `.date` | C# tuple không dùng named members | Đổi thành `.Item1`/`.Item2` |
| 3 | `Timer.Restart()` not found | WinForms Timer không có Restart | Đổi thành `Stop(); Start()` |
| 4 | Lambda syntax CS1026 | Thiếu `{}` trong `Invoke(() => Stop(); Start())` | Thêm braces |
| 5 | `SplitterDistance` InvalidOperationException | Set trong constructor/Load trước khi layout xong | Chuyển sang `Shown` event + try-catch |
| 6 | Column headers không hiển thị | `pDB` panel (DockStyle.Top) che `dgvDB` (DockStyle.Fill) | Xóa pDB, đưa nút lên toolbar chính |
| 7 | Buttons tràn khỏi toolbar | Quá nhiều button trên 1 hàng với form 1200px | Tách toolbar thành 2 hàng (height=82px) |
| 8 | `lblWatch` được add 2 lần | Tạo trong BuildUI VÀ StartWatcher đều gọi Controls.Add | Khởi tạo trong BuildUI, StartWatcher chỉ update `.Text` |

---

## ⚠️ Trạng thái hiện tại

- **Build:** ✅ Thành công (MSBuild, không có lỗi CS)
- **Runtime SplitterDistance:** ⏳ Chờ user xác nhận fix `Shown` event đã giải quyết chưa
- **Column headers:** ⏳ Chờ user xác nhận hiển thị đúng sau khi xóa `pDB`

### Nếu SplitterDistance vẫn lỗi:
Fallback: xóa hoàn toàn dòng set `SplitterDistance` → để SplitContainer dùng 50/50 mặc định

---

## 📝 Ghi chú kỹ thuật quan trọng

1. **WinForms Dock LIFO**: `Controls.Add` theo thứ tự ngược với ý muốn hiển thị. Fill phải add TRƯỚC Bottom/Top.
2. **SplitContainer.SplitterDistance**: PHẢI set trong `Shown` event, không phải `Load` hay constructor. Giá trị phải nằm trong `[Panel1MinSize, Width - Panel2MinSize]`.
3. **MSBuild path**: `C:\Program Files\Microsoft Visual Studio\18\Professional\MSBuild\Current\Bin\MSBuild.exe` — không dùng `dotnet build` cho project này (MSB4803 error).
4. **EPPlus**: `ExcelPackage.LicenseContext = LicenseContext.NonCommercial` phải set trước khi dùng.
5. **Tuple named members**: C# `(string path, DateTime date)` chỉ dùng được named members trong C# 7.0+ với ValueTuple. Khi có vấn đề, dùng `.Item1`/`.Item2` an toàn hơn.
