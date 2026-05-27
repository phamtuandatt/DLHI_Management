using System;
using System.Collections.Generic;
using System.Data;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;

namespace MPR_Managerment.Services
{
    /// <summary>
    /// Dịch vụ AI tìm kiếm toàn hệ thống — dùng Local LLM Proxy (OpenAI-compatible) + truy vấn DB thông minh.
    /// </summary>
    public class AISearchService
    {
        // ── Cấu hình 9Router Proxy (OpenAI-compatible) ────────────────────
        // Proxy chạy local tại: http://localhost:20128/v1/chat/completions
        private const string ROUTER_API_URL = "http://localhost:20128/v1/chat/completions";
        private const string ROUTER_MODEL = "ForVSCode";

        // API key cho 9Router Proxy — để trống nếu proxy không yêu cầu xác thực
        private const string ROUTER_API_KEY = "sk-be73dd5e6a579e85-suubtl-b7961654";

        private static readonly HttpClient _http = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(120)
        };

        public AISearchService(string apiKey = "") { }

        /// <summary>Warm-up không cần thiết với API online.</summary>
        public Task WarmUpAsync() => Task.CompletedTask;

        // ── Lịch sử hội thoại (multi-turn) ────────────────────────────────
        private readonly List<(string role, string text)> _history = new();

        public void ClearHistory() => _history.Clear();

        // ── Schema DB thực tế — tên bảng và cột chính xác từ DB ──────────
        private const string DB_SCHEMA = @"
Database schema (SQL Server) của phần mềm quản lý vật tư MPR_Management:

=== BẢNG CHÍNH ===

MPR_Header(MPR_ID int PK, MPR_No nvarchar(50), Project_Name nvarchar(255), Project_Code nvarchar(50),
  Department nvarchar(100), Requestor nvarchar(100), Rev nvarchar(50) [VARCHAR — dùng TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT) khi so sánh số],
  Required_Date date, Status nvarchar(50), Notes nvarchar(500), Is_Latest bit, Created_Date datetime)

MPR_Details(Detail_ID int PK, MPR_ID int FK→MPR_Header, Item_No nvarchar(50) [VARCHAR],
  item_name nvarchar(500), Description nvarchar(500) [tên cột: Description_Line1/Description_Line2],
  Material nvarchar(200), Thickness_mm decimal, Depth_mm decimal, C_Width_mm decimal,
  D_Web_mm decimal, E_Flange_mm decimal, F_Length_mm decimal, Usage_Location nvarchar(200),
  MPS_Info nvarchar(100), REV nvarchar(50), DWG_BOQ_Receive_Date date, Issue_Date date,
  UNIT nvarchar(50), Qty_Per_Sheet int, Weight_kg decimal, Remarks nvarchar(500),
  Is_Deleted bit [0=còn hiệu lực])

PO_head(PO_ID int PK, PONo nvarchar(50) UNIQUE, Project_Name nvarchar(255), MPR_No nvarchar(50),
  Supplier_ID int FK→Suppliers, PO_Date date, Total_Amount decimal, Status nvarchar(50),
  Expected_Delivery datetime, Payment_Term nvarchar(200), ProjectCode varchar(50),
  Notes nvarchar(500), Created_Date datetime, Created_By nvarchar(100))

PO_Detail(PO_Detail_ID int PK, PO_ID int FK→PO_head, Item_No int, item_name nvarchar(500),
  Material nvarchar(200), Qty_Per_Sheet decimal, UNIT nvarchar(50), Weight_kg decimal,
  Price decimal, Amount decimal, VAT decimal, Received int, Received_Qty decimal,
  Status_Delivery bit, MPR_Detail_ID int FK→MPR_Details, Supplier_ID int,
  RequestDay date, DeliveryLocation nvarchar(500))

Suppliers(Supplier_ID int PK, Company_Name nvarchar(255), Short_Name nvarchar(100),
  Supplier_Type nvarchar(100), Email nvarchar(255), Contact_Person nvarchar(200),
  Contact_Phone nvarchar(50), Tax_Code nvarchar(50), IsActive bit)

RIR_head(RIR_ID int PK, RIR_No nvarchar(50) UNIQUE, Issue_Date date, Project_Name nvarchar(255),
  PONo nvarchar(50) FK→PO_head, MPR_No nvarchar(50), Status nvarchar(50),
  Created_Date datetime, Created_By nvarchar(100))

RIR_detail(RIR_Detail_ID int PK, RIR_ID int FK→RIR_head, PO_Detail_ID int,
  Item_No int, item_name nvarchar(500), Material nvarchar(200), UNIT nvarchar(50),
  Qty_Required decimal, Qty_Received decimal, Inspect_Result nvarchar(20), Remarks nvarchar(max))

PO_PrintRequestHistory(Print_ID int PK, PONo nvarchar(100), Project_Name nvarchar(200),
  Dot_TT int, Dot_Label nvarchar(10), Amount_Net decimal, Amount_VAT decimal,
  Amount_Total decimal, Printed_By nvarchar(100), Printed_Date datetime,
  Supplier_Short nvarchar(100))

PO_PaymentProgress(Progress_ID int PK, Print_ID int FK→PO_PrintRequestHistory,
  PONo nvarchar(100), PR_Status nvarchar(50), PR_Paid bit, Amount_Total decimal,
  Dot_TT nvarchar(50), EC_Status nvarchar(50), PR_Note nvarchar(500), Updated_At datetime)

PO_HistoryPaid(HP_ID int PK, Print_ID int, PONo nvarchar(100), Amount_Total decimal,
  PR_Note nvarchar(500))

PO_DeliveryTracking(TrackID int PK, PONo nvarchar(100), ExpDelivery date,
  Status nvarchar(50), GhiChu nvarchar(500), ReceiverNote nvarchar(500), Created_Date datetime)

Warehouse_Import(Import_ID int PK, Import_No nvarchar(50), Import_Date date,
  PO_ID int, PO_Detail_ID int, RIR_ID int, Item_Name nvarchar(500),
  Material nvarchar(200), UNIT nvarchar(50), Qty_Import decimal, Weight_kg decimal,
  Project_Code nvarchar(50), Location nvarchar(200), Created_By nvarchar(100), Created_Date datetime)

ProjectInfo(Project_Code varchar(50) PK, Project_Name nvarchar(255), ...)

=== VIEW HỮU ÍCH ===
vw_PO_FullInfo          — thông tin PO đầy đủ kèm Supplier
vw_MPR_Full_Info        — thông tin MPR đầy đủ
vw_Supplier_FullInfo    — Supplier kèm contacts, certificates, bank accounts
vw_PO_Payment_Summary   — tổng hợp thanh toán PO
vw_Supplier_Debt_Summary — công nợ nhà cung cấp
vw_Warehouse_Stock_V2   — tồn kho hiện tại

=== LƯU Ý QUAN TRỌNG ===
- NCC = Nhà cung cấp = bảng [Suppliers] (KHÔNG phải 'Supplier')
- Bảng MPR: Is_Latest=1 là bản mới nhất; Is_Deleted=0 trong MPR_Details là còn hiệu lực
- Rev và Item_No trong MPR là VARCHAR → dùng TRY_CAST(TRY_CAST(col AS DECIMAL(10,2)) AS INT)
- JOIN PO với NCC: PO_head.Supplier_ID = Suppliers.Supplier_ID
- Lịch sử thanh toán: PO_PrintRequestHistory JOIN PO_PaymentProgress ON Print_ID
- Tiến độ giao hàng: PO_DeliveryTracking.PONo = PO_head.PONo
";

        // ════════════════════════════════════════════════════════════════════
        // BỘ NHỚ AI — lưu quy tắc truy xuất dữ liệu vào bảng AI_Memory (DB)
        // ════════════════════════════════════════════════════════════════════

        /// <summary>Lấy tất cả quy tắc đang hoạt động từ DB.</summary>
        // ── Cache memory để tránh query DB mỗi lần gọi prompt ────────────
        private static List<(int Id, string Rule)> _memoryCache = null;
        private static DateTime _memoryCacheTime = DateTime.MinValue;
        private static readonly TimeSpan CACHE_TTL = TimeSpan.FromMinutes(5);

        public static List<(int Id, string Rule)> GetMemories(bool forceRefresh = false)
        {
            // Trả cache nếu còn mới
            if (!forceRefresh && _memoryCache != null
                && DateTime.Now - _memoryCacheTime < CACHE_TTL)
                return _memoryCache;

            var result = new List<(int, string)>();
            try
            {
                // Dùng connection string riêng với MARS=true
                var builder = new SqlConnectionStringBuilder(
                    DatabaseHelper.GetConnection().ConnectionString)
                {
                    MultipleActiveResultSets = true,
                    ApplicationName = "MPR_AI_Memory"
                };
                using var conn = new SqlConnection(builder.ConnectionString);
                conn.Open();
                var dt = new System.Data.DataTable();
                new SqlDataAdapter(
                    new SqlCommand(
                        "SELECT Memory_ID, Rule_Text FROM AI_Memory " +
                        "WHERE Is_Active = 1 ORDER BY Memory_ID", conn)).Fill(dt);
                foreach (System.Data.DataRow r in dt.Rows)
                    result.Add((Convert.ToInt32(r["Memory_ID"]), r["Rule_Text"].ToString()));

                // Cập nhật cache
                _memoryCache = result;
                _memoryCacheTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AI_Memory] GetMemories lỗi: {ex.Message}");
                // Nếu bảng chưa tạo → trả rỗng, không crash
            }
            return result;
        }

        // Invalidate cache sau mỗi thao tác thêm/xóa
        private static void InvalidateMemoryCache() => _memoryCache = null;

        /// <summary>Thêm quy tắc mới vào DB.</summary>
        public static string AddMemory(string rule, string createdBy = "User")
        {
            rule = rule?.Trim() ?? "";
            if (string.IsNullOrEmpty(rule)) return "⚠️ Quy tắc trống, không lưu.";
            try
            {
                var builder = new SqlConnectionStringBuilder(
                    DatabaseHelper.GetConnection().ConnectionString)
                { MultipleActiveResultSets = true, ApplicationName = "MPR_AI_Memory" };
                using var conn = new SqlConnection(builder.ConnectionString);
                conn.Open();

                // Kiểm tra trùng
                var chk = new SqlCommand(
                    "SELECT COUNT(*) FROM AI_Memory WHERE Rule_Text = @r AND Is_Active = 1", conn);
                chk.Parameters.AddWithValue("@r", rule);
                if (Convert.ToInt32(chk.ExecuteScalar()) > 0)
                    return "✅ Quy tắc đã tồn tại trong bộ nhớ.";

                var ins = new SqlCommand(
                    "INSERT INTO AI_Memory (Rule_Text, Created_By) VALUES (@r, @u)", conn);
                ins.Parameters.AddWithValue("@r", rule);
                ins.Parameters.AddWithValue("@u", createdBy);
                ins.ExecuteNonQuery();

                InvalidateMemoryCache(); // Xóa cache để lần sau lấy mới
                return $"✅ Đã ghi nhớ: \"{rule}\"";
            }
            catch (Exception ex) { return $"⚠️ Lỗi lưu bộ nhớ: {ex.Message}"; }
        }

        /// <summary>Vô hiệu hóa quy tắc theo ID.</summary>
        public static string RemoveMemory(int memoryId)
        {
            try
            {
                var builder = new SqlConnectionStringBuilder(
                    DatabaseHelper.GetConnection().ConnectionString)
                { MultipleActiveResultSets = true, ApplicationName = "MPR_AI_Memory" };
                using var conn = new SqlConnection(builder.ConnectionString);
                conn.Open();

                var get = new SqlCommand(
                    "SELECT Rule_Text FROM AI_Memory WHERE Memory_ID = @id", conn);
                get.Parameters.AddWithValue("@id", memoryId);
                string rule = get.ExecuteScalar()?.ToString() ?? "";

                var del = new SqlCommand(
                    "UPDATE AI_Memory SET Is_Active = 0 WHERE Memory_ID = @id", conn);
                del.Parameters.AddWithValue("@id", memoryId);
                del.ExecuteNonQuery();

                InvalidateMemoryCache();
                return $"✅ Đã xóa: \"{rule}\"";
            }
            catch (Exception ex) { return $"⚠️ Lỗi xóa: {ex.Message}"; }
        }

        /// <summary>Vô hiệu hóa tất cả quy tắc.</summary>
        public static string ClearMemories()
        {
            try
            {
                var builder = new SqlConnectionStringBuilder(
                    DatabaseHelper.GetConnection().ConnectionString)
                { MultipleActiveResultSets = true, ApplicationName = "MPR_AI_Memory" };
                using var conn = new SqlConnection(builder.ConnectionString);
                conn.Open();
                new SqlCommand("UPDATE AI_Memory SET Is_Active = 0", conn).ExecuteNonQuery();
                InvalidateMemoryCache();
                return "✅ Đã xóa toàn bộ quy tắc.";
            }
            catch (Exception ex) { return $"⚠️ Lỗi: {ex.Message}"; }
        }

        /// <summary>Inject quy tắc vào prompt AI.</summary>
        private static string BuildMemoryContext()
        {
            var mems = GetMemories();
            if (mems.Count == 0) return "";
            var sb = new StringBuilder();
            sb.AppendLine("=== QUY TẮC TRUY XUẤT DỮ LIỆU (bắt buộc tuân theo khi viết SQL) ===");
            foreach (var (_, rule) in mems)
                sb.AppendLine($"- {rule}");
            sb.AppendLine("==========================================================");
            return sb.ToString();
        }


        private const string DB_SCHEMA_SHORT = @"
Bảng SQL Server (chỉ SELECT):
MPR_Header: MPR_ID, MPR_No, Project_Name, Project_Code, Rev(varchar), Required_Date, Status, Is_Latest, Created_Date
MPR_Details: Detail_ID, MPR_ID, Item_No(varchar), item_name, Material, Thickness_mm, Depth_mm, C_Width_mm, D_Web_mm, E_Flange_mm, F_Length_mm, UNIT, Qty_Per_Sheet, Weight_kg, REV, Is_Deleted, Remarks
PO_head: PO_ID, PONo, MPR_No, Supplier_ID, PO_Date, Total_Amount, Status, Expected_Delivery, Payment_Term, Project_Name
PO_Detail: PO_Detail_ID, PO_ID, item_name, Qty_Per_Sheet, Weight_kg, Price, Amount, VAT, Received_Qty, MPR_Detail_ID, Status_Delivery, RequestDay
Suppliers: Supplier_ID, Company_Name, Short_Name, Contact_Person, Contact_Phone, IsActive
RIR_head: RIR_ID, RIR_No, Issue_Date, PONo, MPR_No, Status, Project_Name
RIR_detail: RIR_Detail_ID, RIR_ID, item_name, Qty_Required, Qty_Received, Inspect_Result
PO_PrintRequestHistory: Print_ID, PONo, Project_Name, Amount_Net, Amount_VAT, Amount_Total, Printed_By, Printed_Date, Supplier_Short
PO_PaymentProgress: Progress_ID, Print_ID, PONo, PR_Status, PR_Paid, Amount_Total, EC_Status, Updated_At
PO_DeliveryTracking: TrackID, PONo, ExpDelivery, Status, GhiChu
Warehouse_Import: Import_ID, Import_Date, PO_ID, Item_Name, Qty_Import, Weight_kg, Project_Code
Views: vw_PO_FullInfo, vw_MPR_Full_Info, vw_Supplier_FullInfo, vw_PO_Payment_Summary, vw_Supplier_Debt_Summary

⚠ TÊN CỘT QUAN TRỌNG — PHẢI DÙNG ĐÚNG:
- PO_head: cột số PO là [PONo] KHÔNG phải PO_No hay PO_Number
- RIR_head: cột số RIR là [RIR_No] KHÔNG phải RIR_Number
- MPR_Details: cột tên vật tư là [item_name] KHÔNG phải Item_Name
- PO_head: cột ngày giao là [Expected_Delivery] KHÔNG phải Delivery_Date
- Suppliers: KHÔNG có bảng Supplier (phải có chữ 's' cuối)
- Is_Latest=1 → MPR bản mới nhất; Is_Deleted=0 → MPR_Details còn hiệu lực
- JOIN NCC: PO_head.Supplier_ID = Suppliers.Supplier_ID
- Lịch sử thanh toán: PO_PrintRequestHistory JOIN PO_PaymentProgress ON Print_ID
- Rev/Item_No là VARCHAR → dùng TRY_CAST(TRY_CAST(col AS DECIMAL(10,2)) AS INT)
";


        // ── Hàm chính: 1 lần gọi Gemini duy nhất ─────────────────────────
        // Luồng thông minh:
        //   Câu thường (chào hỏi, hỏi chung) → Gemini trả lời thẳng, 0 truy vấn DB
        //   Câu hỏi dữ liệu → Gemini sinh SQL → chạy DB → Gemini trả lời có dữ liệu
        public async Task<string> AskAsync(string userQuestion,
            Action<string> onChunk = null)
        {
            string historyCtx = BuildHistoryContext(4);

            // ── Prompt all-in-one: Ollama quyết định có cần DB không ──────
            // Dùng schema rút gọn để giảm token, tăng tốc độ phản hồi
            string prompt = $@"Bạn là trợ lý AI của phần mềm quản lý vật tư MPR_Management.
Trả lời tiếng Việt. Trò chuyện bình thường VÀ tra cứu DB khi cần.

{DB_SCHEMA_SHORT}
{BuildMemoryContext()}
Lịch sử:{historyCtx}

Câu hỏi: ""{userQuestion}""

Trả về JSON (không markdown), 1 trong 3 dạng:

1. Cần DB, KHÔNG xuất Excel:
{{""need_sql"":true,""export_excel"":false,""sql"":""SELECT TOP 100 ...""}}

2. Cần DB, CÓ xuất Excel (khi user dùng từ: xuất, export, tạo file, báo cáo, danh sách, tổng hợp):
{{""need_sql"":true,""export_excel"":true,""report_name"":""Tên báo cáo ngắn gọn"",""sql"":""SELECT TOP 5000 ...""}}

3. Không cần DB:
{{""need_sql"":false,""export_excel"":false,""answer"":""trả lời ngắn gọn""}}

Quy tắc:
- SQL chỉ SELECT, Is_Latest=1, Is_Deleted=0, không giới hạn TOP khi xuất Excel.
- export_excel=true khi câu hỏi có: xuất/export/tạo file/báo cáo/danh sách/tổng hợp/thống kê.
- Yêu cầu xóa/sửa: need_sql=false, giải thích AI chỉ đọc.

JSON:";

            // Gọi Gemini 1 lần để lấy intent
            string raw = await CallRouterAsync(prompt, temperature: 0.1f);

            // Parse JSON response
            raw = raw.Trim().TrimStart('`').TrimEnd('`');
            if (raw.StartsWith("json", StringComparison.OrdinalIgnoreCase))
                raw = raw.Substring(4).Trim();

            bool needSql = false;
            bool exportExcel = false;
            string sql = "";
            string reportName = "";
            string directAnswer = "";

            try
            {
                using var doc = JsonDocument.Parse(raw);
                var root = doc.RootElement;
                needSql = root.TryGetProperty("need_sql", out var ns) && ns.GetBoolean();
                exportExcel = root.TryGetProperty("export_excel", out var ex) && ex.GetBoolean();
                reportName = root.TryGetProperty("report_name", out var rn) ? rn.GetString() ?? "" : userQuestion;

                if (needSql)
                    sql = root.TryGetProperty("sql", out var s) ? s.GetString() ?? "" : "";
                else
                    directAnswer = root.TryGetProperty("answer", out var a)
                        ? a.GetString() ?? "" : raw;
            }
            catch
            {
                directAnswer = raw;
            }

            string answer;

            if (!needSql || string.IsNullOrWhiteSpace(sql))
            {
                answer = string.IsNullOrEmpty(directAnswer)
                    ? "Xin lỗi, tôi chưa hiểu rõ câu hỏi. Bạn có thể hỏi lại được không?"
                    : directAnswer;
                onChunk?.Invoke(answer);
            }
            else
            {
                // Cần DB → chạy SQL
                var (dbContext, dt) = await RunSQLWithTableAsync(sql);

                if (dbContext.StartsWith("[Lỗi truy vấn DB:"))
                {
                    // Tự động sửa SQL khi gặp lỗi tên cột
                    string fixedSql = await FixSQLAsync(sql, dbContext);
                    if (!string.IsNullOrEmpty(fixedSql) && fixedSql != sql)
                    {
                        var (dbContext2, dt2) = await RunSQLWithTableAsync(fixedSql);
                        dbContext = dbContext2;
                        dt = dt2;
                    }
                }

                if (exportExcel && dt != null && dt.Rows.Count > 0)
                {
                    // ── Tự động xuất Excel ngay, không cần user bấm nút ──
                    string filePath = ExportToExcelFile(dt, reportName, userQuestion);
                    if (filePath != null)
                    {
                        answer = $"✅ Đã tạo báo cáo Excel: **{System.IO.Path.GetFileName(filePath)}**\n" +
                                 $"📊 {dt.Rows.Count} dòng dữ liệu • {dt.Columns.Count} cột\n" +
                                 $"📁 {filePath}";
                        onChunk?.Invoke(answer);
                        // Mở file tự động
                        _pendingExcelPath = filePath;
                    }
                    else
                    {
                        answer = "[Lỗi tạo file Excel]";
                        onChunk?.Invoke(answer);
                    }
                }
                else
                {
                    answer = await GenerateAnswerAsync(userQuestion, dbContext, onChunk);
                }
            }

            // Lưu lịch sử in-memory
            _history.Add(("user", userQuestion));
            _history.Add(("model", answer));

            return answer;
        }

        // ── Path file Excel vừa tạo — frmAIChat đọc để mở file ──────────
        public string _pendingExcelPath = null;

        /// <summary>Xuất DataTable ra file Excel với format đẹp.</summary>
        private string ExportToExcelFile(DataTable dt, string reportName, string question)
        {
            try
            {
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using var pkg = new OfficeOpenXml.ExcelPackage();
                var ws = pkg.Workbook.Worksheets.Add(reportName.Length > 30
                    ? reportName.Substring(0, 30) : reportName);

                // ── Tiêu đề báo cáo ──
                ws.Cells[1, 1].Value = reportName;
                ws.Cells[1, 1, 1, dt.Columns.Count].Merge = true;
                ws.Cells[1, 1].Style.Font.Bold = true;
                ws.Cells[1, 1].Style.Font.Size = 14;
                ws.Cells[1, 1].Style.HorizontalAlignment =
                    OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[1, 1].Style.Fill.PatternType =
                    OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells[1, 1].Style.Fill.BackgroundColor
                    .SetColor(System.Drawing.Color.FromArgb(0, 70, 127));
                ws.Cells[1, 1].Style.Font.Color
                    .SetColor(System.Drawing.Color.White);

                // ── Dòng thông tin xuất ──
                ws.Cells[2, 1].Value =
                    $"Câu hỏi: {question}   |   Xuất lúc: {DateTime.Now:dd/MM/yyyy HH:mm}   |   Tổng: {dt.Rows.Count} dòng";
                ws.Cells[2, 1, 2, dt.Columns.Count].Merge = true;
                ws.Cells[2, 1].Style.Font.Italic = true;
                ws.Cells[2, 1].Style.Font.Size = 9;
                ws.Cells[2, 1].Style.Font.Color
                    .SetColor(System.Drawing.Color.FromArgb(80, 80, 80));

                // ── Header cột ──
                var headerColor = System.Drawing.Color.FromArgb(0, 120, 212);
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    var cell = ws.Cells[4, c + 1];
                    cell.Value = dt.Columns[c].ColumnName;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(headerColor);
                    cell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    cell.Style.HorizontalAlignment =
                        OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    cell.Style.Border.Bottom.Style =
                        OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                }

                // ── Dữ liệu ──
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        var cell = ws.Cells[r + 5, c + 1];
                        var val = dt.Rows[r][c];
                        cell.Value = val == DBNull.Value ? "" : val;

                        // Zebra stripe
                        if (r % 2 == 1)
                        {
                            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor
                                .SetColor(System.Drawing.Color.FromArgb(240, 246, 255));
                        }
                    }
                }

                // ── Border toàn bảng ──
                if (dt.Rows.Count > 0)
                {
                    var tableRange = ws.Cells[4, 1, dt.Rows.Count + 4, dt.Columns.Count];
                    tableRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                ws.Cells[ws.Dimension.Address].AutoFitColumns(8, 60);

                // ── Freeze panes (cố định header) ──
                ws.View.FreezePanes(5, 1);

                // ── Lưu file ──
                string safeFileName = string.Concat(reportName
                    .Replace("/", "-").Replace("\\", "-")
                    .Replace(":", "").Replace("*", "").Replace("?", "")
                    .Replace("\"", "").Replace("<", "").Replace(">", "")
                    .Replace("|", "").Take(50));
                string path = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    $"AI_{safeFileName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

                pkg.SaveAs(new System.IO.FileInfo(path));
                _lastQueryResult = null; // reset sau khi xuất tự động
                return path;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ExportToExcel error: " + ex.Message);
                return null;
            }
        }

        // ── Placeholder để không break code cũ ───────────────────────────
        private async Task<string> AskAsync_unused() => "";

        // ── Tự động sửa SQL khi gặp lỗi tên cột ─────────────────────────
        private async Task<string> FixSQLAsync(string badSql, string errorMsg)
        {
            try
            {
                string fixPrompt = $@"SQL sau gặp lỗi: {errorMsg}

SQL lỗi:
{badSql}

Schema đúng:
{DB_SCHEMA_SHORT}

Hãy sửa lại SQL cho đúng tên bảng/cột. Chỉ trả về SQL đã sửa, không giải thích.
SQL đã sửa:";

                return await CallRouterAsync(fixPrompt, temperature: 0.1f);
            }
            catch { return badSql; }
        }

        // ── Bước 2: Chạy SQL an toàn ──────────────────────────────────────
        private async Task<string> RunSQLAsync(string sql)
        {
            return await Task.Run(() =>
            {
                try
                {
                    string err = ValidateSQL(sql);
                    if (err != null) return err;

                    // Dùng connection AI riêng + SqlDataAdapter để load toàn bộ vào memory
                    // trước khi đóng — tránh lỗi "open DataReader"
                    using var conn = CreateAIConnection();
                    conn.Open();
                    var dt = new DataTable();
                    using (var adapter = new SqlDataAdapter(
                        new SqlCommand(sql.Trim(), conn) { CommandTimeout = 60 }))
                        adapter.Fill(dt);

                    if (dt.Rows.Count == 0)
                        return "[Không tìm thấy dữ liệu phù hợp]";
                    return DataTableToText(dt, maxRows: 150);
                }
                catch (Exception ex)
                {
                    return $"[Lỗi truy vấn DB: {ex.Message}]";
                }
            });
        }

        // ── Bước 3: Gemini tổng hợp câu trả lời ──────────────────────────
        private async Task<string> GenerateAnswerAsync(string question,
            string dbContext, Action<string> onChunk)
        {
            string historyCtx = BuildHistoryContext(4);

            string dataSection = string.IsNullOrEmpty(dbContext)
                ? ""
                : $"\nDữ liệu lấy từ database:\n{dbContext}\n";

            string prompt = $@"Bạn là trợ lý AI của phần mềm quản lý vật tư MPR_Management. Trả lời tiếng Việt.

{BuildMemoryContext()}
Lịch sử:{historyCtx}
{dataSection}
Câu hỏi: ""{question}""

Hướng dẫn:
- Phân tích dữ liệu DB ở trên và trả lời ngắn gọn, chính xác.
- Số tiền: định dạng ngàn (1.234.567 VNĐ). Ngày: dd/MM/yyyy. KHÔNG bịa số liệu.
- QUAN TRỌNG: Với mỗi MPR_No, PONo, RIR_No, Supplier_ID xuất hiện trong câu trả lời,
  hãy tạo deep-link theo định dạng: [Tên hiển thị](prefix://key)
  Prefix: mpr:// cho MPR_No | po:// cho PONo | rir:// cho RIR_No | ncc:// cho Supplier_ID
  Ví dụ: [DV-FT-2505-MPR-001](mpr://DV-FT-2505-MPR-001) hoặc [PO-2025-001](po://PO-2025-001)
  Người dùng có thể click vào link để mở thẳng record đó trong phần mềm.

Trả lời:";

            // Streaming response nếu có callback
            if (onChunk != null)
                return await CallRouterStreamAsync(prompt, onChunk);
            else
                return await CallRouterAsync(prompt, temperature: 0.7f);
        }

        // ── 9Router API call (non-stream) ────────────────────────────────
        private async Task<string> CallRouterAsync(string prompt, float temperature = 0.5f)
        {
            try
            {
                var body = new
                {
                    model = ROUTER_MODEL,
                    messages = new[] { new { role = "user", content = prompt } },
                    temperature,
                    max_tokens = 2048,
                    stream = false
                };

                string json = JsonSerializer.Serialize(body);
                var request = new HttpRequestMessage(HttpMethod.Post, ROUTER_API_URL)
                {
                    Content = new StringContent(json, Encoding.UTF8, "application/json")
                };
                // Thêm API key vào header nếu có
                if (!string.IsNullOrEmpty(ROUTER_API_KEY))
                    request.Headers.Add("Authorization", $"Bearer {ROUTER_API_KEY}");

                var resp = await _http.SendAsync(request);
                if (!resp.IsSuccessStatusCode)
                {
                    string err = await resp.Content.ReadAsStringAsync();
                    return $"[Lỗi 9Router {resp.StatusCode}: {err.Substring(0, Math.Min(200, err.Length))}]";
                }

                string raw = await resp.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(raw);
                return doc.RootElement
                    .GetProperty("choices")[0]
                    .GetProperty("message")
                    .GetProperty("content")
                    .GetString() ?? "";
            }
            catch (Exception ex)
            {
                return $"[Lỗi 9Router: {ex.Message}]";
            }
        }

        // ── 9Router API call (streaming) ──────────────────────────────────
        // 9Router hỗ trợ SSE streaming — format giống OpenAI-compatible
        private async Task<string> CallRouterStreamAsync(string prompt, Action<string> onChunk)
        {
            var sb = new StringBuilder();
            try
            {
                var body = new
                {
                    model = ROUTER_MODEL,
                    messages = new[] { new { role = "user", content = prompt } },
                    temperature = 0.7,
                    max_tokens = 2048,
                    stream = true
                };

                string json = JsonSerializer.Serialize(body);
                var request = new HttpRequestMessage(HttpMethod.Post, ROUTER_API_URL)
                {
                    Content = new StringContent(json, Encoding.UTF8, "application/json")
                };
                // Thêm API key vào header nếu có
                if (!string.IsNullOrEmpty(ROUTER_API_KEY))
                    request.Headers.Add("Authorization", $"Bearer {ROUTER_API_KEY}");

                using var resp = await _http.SendAsync(request,
                    HttpCompletionOption.ResponseHeadersRead);

                if (!resp.IsSuccessStatusCode)
                {
                    string fallback = await CallRouterAsync(prompt, 0.7f);
                    onChunk?.Invoke(fallback);
                    return fallback;
                }

                using var stream = await resp.Content.ReadAsStreamAsync();
                using var reader = new System.IO.StreamReader(stream);

                while (!reader.EndOfStream)
                {
                    string line = await reader.ReadLineAsync();
                    if (string.IsNullOrEmpty(line) || !line.StartsWith("data: ")) continue;

                    string data = line.Substring(6).Trim();
                    if (data == "[DONE]") break;

                    try
                    {
                        using var doc = JsonDocument.Parse(data);
                        var delta = doc.RootElement
                            .GetProperty("choices")[0]
                            .GetProperty("delta");

                        if (delta.TryGetProperty("content", out var c))
                        {
                            string chunk = c.GetString() ?? "";
                            if (!string.IsNullOrEmpty(chunk))
                            {
                                sb.Append(chunk);
                                onChunk?.Invoke(chunk);
                            }
                        }
                    }
                    catch { /* bỏ qua dòng parse lỗi */ }
                }
            }
            catch (Exception ex)
            {
                string errMsg = $"\n[Lỗi 9Router: {ex.Message}]";
                onChunk?.Invoke(errMsg);
                sb.Append(errMsg);
            }

            return sb.ToString();
        }

        // ── DataTable → chuỗi text gọn cho Gemini ─────────────────────────
        private string DataTableToText(DataTable dt, int maxRows = 150)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Kết quả: {Math.Min(dt.Rows.Count, maxRows)} dòng / " +
                          $"{dt.Columns.Count} cột");

            // Header
            var cols = new List<string>();
            foreach (DataColumn col in dt.Columns) cols.Add(col.ColumnName);
            sb.AppendLine(string.Join(" | ", cols));
            sb.AppendLine(new string('-', Math.Min(cols.Count * 15, 120)));

            // Data rows
            int count = 0;
            foreach (DataRow row in dt.Rows)
            {
                if (count++ >= maxRows) break;
                var vals = new List<string>();
                foreach (DataColumn col in dt.Columns)
                {
                    string v = row[col] == DBNull.Value ? "" : row[col].ToString();
                    // Cắt ngắn giá trị quá dài
                    if (v.Length > 60) v = v.Substring(0, 57) + "...";
                    vals.Add(v);
                }
                sb.AppendLine(string.Join(" | ", vals));
            }

            if (dt.Rows.Count > maxRows)
                sb.AppendLine($"... (còn {dt.Rows.Count - maxRows} dòng nữa, chỉ hiển thị {maxRows})");

            return sb.ToString();
        }

        // ════════════════════════════════════════════════════════════════════
        // CHẾ ĐỘ CHAT TỰ DO — không truy vấn DB, AI trả lời như chatbot thường
        // ════════════════════════════════════════════════════════════════════

        public async Task<string> AskFreeAsync(string userQuestion, Action<string> onChunk = null)
        {
            string historyCtx = BuildHistoryContext(4);
            string memCtx = BuildMemoryContext();

            string prompt = $@"Bạn là trợ lý AI thông minh, hữu ích và thân thiện.
Trả lời tiếng Việt trừ khi người dùng hỏi bằng ngôn ngữ khác.
Bạn có thể: giải thích khái niệm, phân tích vấn đề, viết nội dung,
tư vấn, lập kế hoạch, dịch thuật, và mọi tác vụ của một AI tổng quát.
{memCtx}
Lịch sử:{historyCtx}

Câu hỏi: ""{userQuestion}""

Trả lời:";

            string answer;
            if (onChunk != null)
                answer = await CallRouterStreamAsync(prompt, onChunk);
            else
                answer = await CallRouterAsync(prompt, temperature: 0.8f);

            _history.Add(("user", userQuestion));
            _history.Add(("model", answer));
            return answer;
        }


        // ════════════════════════════════════════════════════════════════════
        // BOOKMARK — lưu câu hỏi yêu thích vào bảng AI_Bookmark (DB)
        // ════════════════════════════════════════════════════════════════════

        private static SqlConnection CreateBookmarkConn()
        {
            var b = new SqlConnectionStringBuilder(
                DatabaseHelper.GetConnection().ConnectionString)
            { MultipleActiveResultSets = true, ApplicationName = "MPR_AI_Bookmark" };
            return new SqlConnection(b.ConnectionString);
        }

        public static List<(int Id, string Question)> GetBookmarks()
        {
            var result = new List<(int, string)>();
            try
            {
                using var conn = CreateBookmarkConn();
                conn.Open();
                var dt = new System.Data.DataTable();
                new SqlDataAdapter(new SqlCommand(
                    "SELECT Bookmark_ID, Question FROM AI_Bookmark " +
                    "WHERE Is_Active = 1 ORDER BY Created_Date DESC", conn)).Fill(dt);
                foreach (System.Data.DataRow r in dt.Rows)
                    result.Add((Convert.ToInt32(r["Bookmark_ID"]), r["Question"].ToString()));
            }
            catch (Exception ex)
            { System.Diagnostics.Debug.WriteLine($"[AI_Bookmark] {ex.Message}"); }
            return result;
        }

        public static string AddBookmark(string question)
        {
            question = question?.Trim() ?? "";
            if (string.IsNullOrEmpty(question)) return "⚠️ Câu hỏi trống.";
            try
            {
                using var conn = CreateBookmarkConn();
                conn.Open();
                var chk = new SqlCommand(
                    "SELECT COUNT(*) FROM AI_Bookmark WHERE Question = @q AND Is_Active = 1", conn);
                chk.Parameters.AddWithValue("@q", question);
                if (Convert.ToInt32(chk.ExecuteScalar()) > 0)
                    return "⭐ Câu hỏi đã có trong bookmark.";
                var ins = new SqlCommand(
                    "INSERT INTO AI_Bookmark (Question) VALUES (@q)", conn);
                ins.Parameters.AddWithValue("@q", question);
                ins.ExecuteNonQuery();
                return $"⭐ Đã lưu bookmark: \"{question}\"";
            }
            catch (Exception ex) { return $"⚠️ Lỗi: {ex.Message}"; }
        }

        public static string RemoveBookmark(int bookmarkId)
        {
            try
            {
                using var conn = CreateBookmarkConn();
                conn.Open();
                var get = new SqlCommand(
                    "SELECT Question FROM AI_Bookmark WHERE Bookmark_ID = @id", conn);
                get.Parameters.AddWithValue("@id", bookmarkId);
                string q = get.ExecuteScalar()?.ToString() ?? "";
                var del = new SqlCommand(
                    "UPDATE AI_Bookmark SET Is_Active = 0 WHERE Bookmark_ID = @id", conn);
                del.Parameters.AddWithValue("@id", bookmarkId);
                del.ExecuteNonQuery();
                return $"✅ Đã xóa: \"{q}\"";
            }
            catch (Exception ex) { return $"⚠️ Lỗi: {ex.Message}"; }
        }

        public static string ClearBookmarks()
        {
            try
            {
                using var conn = CreateBookmarkConn();
                conn.Open();
                new SqlCommand("UPDATE AI_Bookmark SET Is_Active = 0", conn).ExecuteNonQuery();
                return "✅ Đã xóa tất cả bookmark.";
            }
            catch (Exception ex) { return $"⚠️ Lỗi: {ex.Message}"; }
        }

        private string BuildHistoryContext(int lastN)
        {
            if (_history.Count == 0) return "(chưa có lịch sử)";
            int start = Math.Max(0, _history.Count - lastN * 2);
            var sb = new StringBuilder();
            for (int i = start; i < _history.Count; i++)
            {
                var (role, text) = _history[i];
                string label = role == "user" ? "Người dùng" : "AI";
                string truncated = text.Length > 300
                    ? text.Substring(0, 297) + "..."
                    : text;
                sb.AppendLine($"{label}: {truncated}");
            }
            return sb.ToString();
        }

        // ════════════════════════════════════════════════════════════════════
        // TÍNH NĂNG 1 — Báo cáo tóm tắt khi mở app
        // ════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Tạo báo cáo tóm tắt tình hình hiện tại — gọi khi mở chatbox lần đầu.
        /// Truy vấn trực tiếp DB (không qua Gemini) để tiết kiệm quota.
        /// </summary>
        public async Task<string> GetDailySummaryAsync()
        {
            return await Task.Run(() =>
            {
                try
                {
                    // Dùng connection riêng với MARS=true để chạy nhiều query liên tiếp
                    using var conn = CreateAIConnection();
                    conn.Open();

                    string today = DateTime.Today.ToString("dd/MM/yyyy");
                    var sb = new StringBuilder();
                    sb.AppendLine($"📊 **Tóm tắt hôm nay — {today}**\n");

                    // 1. PO sắp hết hạn giao hàng (7 ngày tới)
                    var cmdPO = new SqlCommand(@"
                        SELECT COUNT(*) FROM PO_head
                        WHERE Expected_Delivery BETWEEN GETDATE() AND DATEADD(DAY,7,GETDATE())
                          AND Status NOT IN ('Completed','Closed','Cancelled')", conn);
                    int poSoon = Convert.ToInt32(cmdPO.ExecuteScalar());

                    // 2. PO đã quá hạn
                    var cmdOverdue = new SqlCommand(@"
                        SELECT COUNT(*) FROM PO_head
                        WHERE Expected_Delivery < GETDATE()
                          AND Status NOT IN ('Completed','Closed','Cancelled')", conn);
                    int poOverdue = Convert.ToInt32(cmdOverdue.ExecuteScalar());

                    // 3. MPR mới nhất chưa có PO
                    var cmdMPR = new SqlCommand(@"
                        SELECT COUNT(*) FROM MPR_Header h
                        WHERE h.Is_Latest = 1
                          AND h.Status NOT IN ('Cancelled','Closed')
                          AND NOT EXISTS (
                              SELECT 1 FROM PO_head p WHERE p.MPR_No = h.MPR_No
                          )", conn);
                    int mprNoPO = Convert.ToInt32(cmdMPR.ExecuteScalar());

                    // 4. PO chờ thanh toán
                    var cmdPay = new SqlCommand(@"
                        SELECT COUNT(DISTINCT PONo) FROM PO_PaymentProgress
                        WHERE PR_Paid = 0 OR PR_Paid IS NULL", conn);
                    int payPending = Convert.ToInt32(cmdPay.ExecuteScalar());

                    // 5. RIR mới trong tuần
                    var cmdRIR = new SqlCommand(@"
                        SELECT COUNT(*) FROM RIR_head
                        WHERE Created_Date >= DATEADD(DAY,-7,GETDATE())", conn);
                    int rirWeek = Convert.ToInt32(cmdRIR.ExecuteScalar());

                    // Format báo cáo
                    sb.AppendLine(poOverdue > 0
                        ? $"🔴 **{poOverdue} PO** đã quá hạn giao hàng"
                        : "✅ Không có PO nào quá hạn");

                    sb.AppendLine(poSoon > 0
                        ? $"⚠️ **{poSoon} PO** sắp đến hạn giao trong 7 ngày tới"
                        : "✅ Không có PO nào sắp đến hạn");

                    sb.AppendLine(mprNoPO > 0
                        ? $"📋 **{mprNoPO} MPR** chưa có PO"
                        : "✅ Tất cả MPR đã có PO");

                    sb.AppendLine(payPending > 0
                        ? $"💰 **{payPending} đợt thanh toán** đang chờ xử lý"
                        : "✅ Không có khoản thanh toán nào chờ");

                    sb.AppendLine($"📦 **{rirWeek} RIR** mới trong 7 ngày qua");

                    sb.AppendLine("\n_Hỏi tôi để xem chi tiết bất kỳ mục nào ở trên._");
                    return sb.ToString();
                }
                catch (Exception ex)
                {
                    return $"[Không thể tải báo cáo: {ex.Message}]";
                }
            });
        }

        // ════════════════════════════════════════════════════════════════════
        // TÍNH NĂNG 3 — Xuất Excel từ kết quả truy vấn
        // Lưu DataTable cuối cùng để frmAIChat có thể gọi ExportToExcel
        // ════════════════════════════════════════════════════════════════════

        private DataTable _lastQueryResult = null;
        public bool HasExportableData => _lastQueryResult != null && _lastQueryResult.Rows.Count > 0;

        /// <summary>Chạy SQL và lưu DataTable để xuất Excel sau.</summary>
        // ── Tạo connection mới độc lập cho AI — tránh conflict với app ───
        // Thêm MARS=True để cho phép nhiều DataReader cùng lúc trên 1 connection
        private SqlConnection CreateAIConnection()
        {
            var builder = new SqlConnectionStringBuilder(
                DatabaseHelper.GetConnection().ConnectionString)
            {
                MultipleActiveResultSets = true,
                // Tạo connection pool riêng cho AI — không tranh chấp với app chính
                ApplicationName = "MPR_AI_Query"
            };
            return new SqlConnection(builder.ConnectionString);
        }

        // ── Validate SQL safety (dùng chung cho 2 hàm RunSQL) ────────────
        private string ValidateSQL(string sql)
        {
            string sqlTrimmed = sql.Trim();
            string sqlUpper = sqlTrimmed.ToUpperInvariant();

            if (!sqlUpper.StartsWith("SELECT") && !sqlUpper.StartsWith("WITH"))
                return "[Từ chối: Chỉ cho phép câu lệnh SELECT.]";

            char[] delimiters = { ' ', '\t', '\r', '\n', '(', ')', ';', ',', '\'' };
            string[] dangerousKeywords = {
                "INSERT", "UPDATE", "DELETE", "DROP", "TRUNCATE",
                "ALTER", "CREATE", "RENAME", "REPLACE",
                "EXEC", "EXECUTE", "SP_", "XP_",
                "GRANT", "REVOKE", "DENY",
                "MERGE", "UPSERT",
                "OPENROWSET", "OPENQUERY", "OPENDATASOURCE",
                "BULK INSERT", "SHUTDOWN", "DBCC"
            };
            foreach (string kw in dangerousKeywords)
            {
                int idx = 0;
                while ((idx = sqlUpper.IndexOf(kw, idx, StringComparison.Ordinal)) >= 0)
                {
                    bool beforeOk = idx == 0 || delimiters.Contains(sqlUpper[idx - 1]);
                    int afterIdx = idx + kw.Length;
                    bool afterOk = afterIdx >= sqlUpper.Length
                                 || delimiters.Contains(sqlUpper[afterIdx]);
                    if (beforeOk && afterOk)
                        return $"[Từ chối: Câu lệnh chứa '{kw}' — chỉ được phép SELECT.]";
                    idx += kw.Length;
                }
            }
            // Kiểm tra SELECT INTO
            if (sqlUpper.Contains("SELECT") && sqlUpper.Contains("INTO "))
                return "[Từ chối: Không cho phép SELECT INTO.]";

            return null; // null = hợp lệ
        }

        public async Task<(string text, DataTable dt)> RunSQLWithTableAsync(string sql)
        {
            return await Task.Run(() =>
            {
                try
                {
                    string err = ValidateSQL(sql);
                    if (err != null) return (err, null as DataTable);

                    // Dùng connection MỚI hoàn toàn — không dùng chung với app
                    using var conn = CreateAIConnection();
                    conn.Open();

                    var cmd = new SqlCommand(sql.Trim(), conn)
                    {
                        CommandTimeout = 60
                    };

                    var dt = new DataTable();
                    // Dùng SqlDataAdapter thay vì dt.Load() — load toàn bộ vào memory
                    // trước khi đóng reader, tránh lỗi "open DataReader"
                    using (var adapter = new SqlDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                    // conn tự đóng khi ra khỏi using

                    _lastQueryResult = dt.Rows.Count > 0 ? dt : null;
                    return (DataTableToText(dt, 150), dt);
                }
                catch (Exception ex)
                {
                    return ($"[Lỗi truy vấn DB: {ex.Message}]", null as DataTable);
                }
            });
        }

        /// <summary>Xuất DataTable cuối cùng ra file Excel.</summary>
        public string ExportLastResultToExcel(string question)
        {
            if (_lastQueryResult == null || _lastQueryResult.Rows.Count == 0)
                return null;
            try
            {
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using var pkg = new OfficeOpenXml.ExcelPackage();
                var ws = pkg.Workbook.Worksheets.Add("Kết quả AI");

                // Tiêu đề
                ws.Cells[1, 1].Value = $"Kết quả: {question}";
                ws.Cells[1, 1, 1, _lastQueryResult.Columns.Count].Merge = true;
                ws.Cells[1, 1].Style.Font.Bold = true;
                ws.Cells[1, 1].Style.Font.Size = 12;
                ws.Cells[2, 1].Value = $"Xuất lúc: {DateTime.Now:dd/MM/yyyy HH:mm}";
                ws.Cells[2, 1, 2, _lastQueryResult.Columns.Count].Merge = true;

                // Header
                for (int c = 0; c < _lastQueryResult.Columns.Count; c++)
                {
                    var cell = ws.Cells[4, c + 1];
                    cell.Value = _lastQueryResult.Columns[c].ColumnName;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 120, 212));
                    cell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                }

                // Data
                for (int r = 0; r < _lastQueryResult.Rows.Count; r++)
                    for (int c = 0; c < _lastQueryResult.Columns.Count; c++)
                    {
                        var val = _lastQueryResult.Rows[r][c];
                        ws.Cells[r + 5, c + 1].Value = val == DBNull.Value ? "" : val;
                    }

                ws.Cells[ws.Dimension.Address].AutoFitColumns();

                string path = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    $"AI_Export_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                pkg.SaveAs(new System.IO.FileInfo(path));
                _lastQueryResult = null; // reset sau khi xuất
                return path;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
}