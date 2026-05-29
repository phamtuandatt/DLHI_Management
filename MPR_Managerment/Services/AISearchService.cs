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
            sb.AppendLine("=== QUY TẮC BẮT BUỘC (PHẢI ĐỌC VÀ ÁP DỤNG TRƯỚC KHI TRẢ LỜI) ===");
            sb.AppendLine("TRƯỚC KHI viết SQL hoặc trả lời, bắt buộc phải:");
            sb.AppendLine("1. Đọc từng quy tắc dưới đây.");
            sb.AppendLine("2. Kiểm tra câu hỏi có liên quan đến quy tắc nào không.");
            sb.AppendLine("3. Nếu có → áp dụng quy tắc đó vào SQL/câu trả lời, KHÔNG bỏ qua.");
            sb.AppendLine();
            sb.AppendLine("DANH SÁCH QUY TẮC:");
            int i = 1;
            foreach (var (_, rule) in mems)
                sb.AppendLine($"  [{i++}] {rule}");
            sb.AppendLine("=======================================================");
            return sb.ToString();
        }


        private const string DB_SCHEMA_SHORT = @"
=== BẢNG SQL SERVER (chỉ SELECT) ===

MPR_Header: MPR_ID, MPR_No, Project_Name, Project_Code, Department, Requestor,
  Rev(varchar — dùng TRY_CAST(TRY_CAST(Rev AS DECIMAL(10,2)) AS INT) khi so sánh số),
  Required_Date(date), Status, Is_Latest(bit), Created_Date(datetime), Notes
  ↳ Status hợp lệ: 'Draft' | 'Submitted' | 'Approved' | 'In Progress' | 'Completed' | 'Cancelled' | 'Closed'

MPR_Details: Detail_ID, MPR_ID, Item_No(varchar), item_name, Description_Line1, Description_Line2,
  Material, Thickness_mm, Depth_mm, C_Width_mm, D_Web_mm, E_Flange_mm, F_Length_mm,
  Usage_Location, UNIT, Qty_Per_Sheet(int), Weight_kg, REV(varchar), Is_Deleted(bit),
  DWG_BOQ_Receive_Date, Issue_Date, Remarks
  ↳ Is_Deleted=0 → còn hiệu lực; Item_No là VARCHAR

PO_head: PO_ID, PONo, MPR_No, Supplier_ID, PO_Date(date), Total_Amount(decimal),
  Status, Expected_Delivery(datetime), Payment_Term, Project_Name, ProjectCode,
  Notes, Created_Date, Created_By
  ↳ Status hợp lệ: 'Draft' | 'Approved' | 'Sent' | 'Partial' | 'Completed' | 'Cancelled' | 'Closed'
  ↳ Số PO dùng cột [PONo] (KHÔNG phải PO_No hay PO_Number)

PO_Detail: PO_Detail_ID, PO_ID, Item_No(int), item_name, Material, Qty_Per_Sheet(decimal),
  UNIT, Weight_kg, Price(decimal), Amount(decimal), VAT(decimal),
  Received(int), Received_Qty(decimal), Status_Delivery(bit),
  MPR_Detail_ID, Supplier_ID, RequestDay(date), DeliveryLocation
  ↳ item_name viết thường; Status_Delivery=1 → đã giao đủ

Suppliers: Supplier_ID, Company_Name, Short_Name, Supplier_Type,
  Cert, Email, Contact_Person, Contact_Phone, Company_Address,
  Bank_Account, Bank_Name, Tax_Code, Website, Notes,
  Zalo_Group_ID, IsActive(bit), Created_Date, Created_By
  ↳ Bảng tên [Suppliers] có chữ s (KHÔNG phải Supplier)
  ↳ Tài khoản ngân hàng: Bank_Account + Bank_Name (trực tiếp trong Suppliers)
  ↳ Chứng chỉ: cột [Cert] trong Suppliers

RIR_head: RIR_ID, RIR_No, Issue_Date(date), Project_Name, PONo, MPR_No,
  Status, Created_Date, Created_By
  ↳ Status hợp lệ: 'Draft' | 'Inspecting' | 'Passed' | 'Failed' | 'Completed'
  ↳ Số RIR dùng cột [RIR_No] (KHÔNG phải RIR_Number)

RIR_detail: RIR_Detail_ID, RIR_ID, PO_Detail_ID, Item_No(int), item_name, Material,
  UNIT, Qty_Required(decimal), Qty_Received(decimal), Inspect_Result(nvarchar), Remarks

PO_Payment_Schedule: Schedule_ID, PO_ID, Dot_TT(int), Payment_Type, Pay_Method,
  Percent_TT(decimal), Amount_Plan(decimal), Due_Date(date), Delivery_Ref,
  Description, Status, Created_Date, Created_By
  ↳ Kế hoạch thanh toán từng đợt; JOIN PO_head ON PO_ID
  ↳ Status: 'Pending' | 'Paid' | 'Overdue' | 'Cancelled'

PO_Payment_History: Payment_ID, Schedule_ID, PO_ID, Supplier_ID,
  Payment_Date(date), Amount_Paid(decimal), Payment_Method, Bank_Name,
  Transaction_No, Currency, Exchange_Rate, Notes, Created_By
  ↳ Lịch sử thanh toán thực tế; JOIN PO_Payment_Schedule ON Schedule_ID

PO_PrintRequestHistory: Print_ID, PONo, Project_Name, Dot_TT(int), Dot_Label,
  Amount_Net, Amount_VAT, Amount_Total, Printed_By, Printed_Date, Supplier_Short
PO_PaymentProgress: Progress_ID, Print_ID, PONo, PR_Status, PR_Paid(bit),
  Amount_Total, Dot_TT, EC_Status, PR_Note, Updated_At
PO_DeliveryTracking: TrackID, PONo, ExpDelivery(date), Status, GhiChu, ReceiverNote, Created_Date

Warehouse_Import: Import_ID, Import_No, Import_Date(date), PO_ID, PO_Detail_ID, RIR_ID,
  Item_Name, Material, UNIT, Qty_Import(decimal), Weight_kg,
  Project_Code, Location, Created_By, Created_Date
Warehouse_Export: Export_ID, Export_No, Export_Date(date), Import_ID, Item_Name, Material,
  Size, UNIT, Qty_Export(decimal), Weight_kg, ID_Code, Project_Code,
  WorkorderNo, Export_To, Purpose, Notes, Created_By, Created_Date

ProjectMaterialTransformTransaction:
  New_Value_Location, Old_Value_Location, Item_Name, Size, Number_Tranform
  ↳ Lịch sử chuyển kho giữa các vị trí

ProjectInfo: Project_Code(PK,varchar), Project_Name, WorkorderNo, PO_Link

=== VIEWS (dùng ưu tiên cho báo cáo) ===
vw_PO_FullInfo          — PO + Supplier: PONo, Project_Name, Company_Name, Short_Name, Total_Amount, Status, PO_Date, Expected_Delivery
vw_MPR_Full_Info        — MPR + chi tiết: MPR_No, Project_Name, item_name, Status, Is_Latest, Rev, Required_Date
vw_Supplier_FullInfo    — Supplier + contacts: Supplier_ID, Company_Name, Email, Contact_Person, Bank_Account, Bank_Name, Cert
vw_PO_Payment_Summary   — Thanh toán PO: PO_ID, PONo, Total_Amount, Amount_Paid, Amount_Remaining, PO_Date
vw_Supplier_Debt_Summary — Công nợ NCC: Supplier_ID, Company_Name, Total_Debt, PO_Count
vw_Warehouse_Stock_V2   — Tồn kho: Item_Name, Material, UNIT, Qty_Stock, Project_Code, Location, Import_Date

=== QUY TẮC BẮT BUỘC KHI VIẾT SQL ===
1. PO_head: [PONo]  |  RIR_head: [RIR_No]  |  MPR_Details: [item_name] (chữ thường)
2. Is_Latest=1 → MPR bản mới nhất  |  Is_Deleted=0 → MPR_Details còn hiệu lực
3. JOIN NCC: PO_head.Supplier_ID = Suppliers.Supplier_ID
4. Kế hoạch TT: PO_Payment_Schedule JOIN PO_head ON PO_ID
5. Lịch sử TT: PO_Payment_History JOIN PO_Payment_Schedule ON Schedule_ID
6. Tồn kho: dùng vw_Warehouse_Stock_V2 (Qty_Stock > 0)
7. Rev/Item_No là VARCHAR → TRY_CAST(TRY_CAST(col AS DECIMAL(10,2)) AS INT) khi so sánh số
8. KHÔNG dùng bảng Supplier_Certificates, Supplier_Bank_Accounts (không tồn tại)
9. Ngày giao PO: [Expected_Delivery]  |  Ngày đến hạn TT: [Due_Date]
";


        // ── Hàm chính: 1 lần gọi Gemini duy nhất ─────────────────────────
        // Luồng thông minh:
        //   Câu thường (chào hỏi, hỏi chung) → Gemini trả lời thẳng, 0 truy vấn DB
        //   Câu hỏi dữ liệu → Gemini sinh SQL → chạy DB → Gemini trả lời có dữ liệu
        public async Task<string> AskAsync(string userQuestion,
            Action<string> onChunk = null)
        {
            string historyCtx = BuildHistoryContext(6);

            // ── Prompt all-in-one: Ollama quyết định có cần DB không ──────
            // Dùng schema rút gọn để giảm token, tăng tốc độ phản hồi
            string prompt = $@"Bạn là trợ lý AI của phần mềm quản lý vật tư MPR_Management.
Trả lời tiếng Việt. Trò chuyện bình thường VÀ tra cứu DB khi cần.

{BuildMemoryContext()}
{DB_SCHEMA_SHORT}
Lịch sử:{historyCtx}

Câu hỏi: ""{userQuestion}""

Trả về JSON (không markdown), 1 trong các dạng sau:

1. Cần DB, KHÔNG xuất Excel:
{{""need_sql"":true,""export_excel"":false,""sql"":""SELECT TOP 100 ...""}}

2. Cần DB, CÓ xuất Excel (khi user dùng từ: xuất, export, tạo file, báo cáo, danh sách, tổng hợp):
{{""need_sql"":true,""export_excel"":true,""report_name"":""Tên báo cáo ngắn gọn"",""sql"":""SELECT TOP 5000 ...""}}

3. Không cần DB:
{{""need_sql"":false,""export_excel"":false,""answer"":""trả lời ngắn gọn""}}

LƯU Ý VỀ GHI NHỚ TỰ ĐỘNG (AUTO-LEARN — ÁP DỤNG TIÊU CHÍ NGHIÊM NGẶT):
Chỉ thêm trường ""new_memory"" khi người dùng YÊU CẦU RÕ RÀNG ghi nhớ (""hãy nhớ..."", ""từ nay..."", ""lưu lại..."") VÀ nội dung cần nhớ phải là:
  ✅ Tên cột/bảng thực tế khác với schema (vd: ""cột ngày giao dùng Ship_Date không phải Expected_Delivery"")
  ✅ Quy tắc nghiệp vụ cụ thể (vd: ""Dự án DVFT luôn tính theo đơn vị kg không phải tấn"")
  ✅ Mapping tên tiếng Việt → tên DB (vd: ""'nhà thầu' = bảng Suppliers"")
KHÔNG ghi nhớ: chào hỏi, câu hỏi thông thường, yêu cầu không liên quan đến dữ liệu, thông tin mơ hồ.
Nếu đủ tiêu chí: {{""need_sql"":false,""export_excel"":false,""answer"":""✅ Đã ghi nhớ."",""new_memory"":""<quy tắc ngắn gọn cụ thể>""}}

Quy tắc:
- SQL chỉ SELECT, Is_Latest=1, Is_Deleted=0, không giới hạn TOP khi xuất Excel.
- export_excel=true khi câu hỏi có: xuất/export/tạo file/báo cáo/danh sách/tổng hợp/thống kê.
- Yêu cầu xóa/sửa: need_sql=false, giải thích AI chỉ đọc.
- Nếu câu hỏi chứa tên bảng/cột hoặc mô tả truy vấn dữ liệu → LUÔN need_sql=true, sinh SQL đầy đủ.
- KHÔNG trả về need_sql=false với answer trống — nếu không chắc chắn, cứ sinh SQL thử.

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

                if (root.TryGetProperty("new_memory", out var nm))
                {
                    string newMem = nm.GetString() ?? "";
                    if (!string.IsNullOrWhiteSpace(newMem))
                    {
                        AddMemory(newMem, "AI_AutoLearn");
                    }
                }

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

            // ── Retry: AI trả về need_sql=false với answer trống → thử sinh SQL thẳng ──
            // Xảy ra khi câu hỏi quá dài/phức tạp, AI bị nhầm sang "chat mode"
            if (!needSql && string.IsNullOrWhiteSpace(directAnswer))
            {
                string retrySql = await ForceSQLGenerationAsync(userQuestion);
                if (!string.IsNullOrEmpty(retrySql))
                {
                    needSql = true;
                    sql = retrySql;
                }
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
                    // Tự động sửa SQL — thử tối đa 3 lần, mỗi lần tích lũy lỗi cũ
                    string currentSql = sql;
                    string accumulatedErrors = dbContext;
                    for (int fixAttempt = 1; fixAttempt <= 3; fixAttempt++)
                    {
                        string fixedSql = await FixSQLAsync(currentSql, accumulatedErrors, fixAttempt);
                        if (string.IsNullOrEmpty(fixedSql) || fixedSql == currentSql) break;

                        var (dbContext2, dt2) = await RunSQLWithTableAsync(fixedSql);
                        if (!dbContext2.StartsWith("[Lỗi truy vấn DB:"))
                        {
                            // Sửa thành công
                            dbContext = dbContext2;
                            dt = dt2;
                            break;
                        }
                        // Lần sau sẽ nhận cả lỗi cũ lẫn lỗi mới để AI có thêm context
                        accumulatedErrors = $"Lần {fixAttempt}: {accumulatedErrors}\nLần {fixAttempt + 1}: {dbContext2}";
                        currentSql = fixedSql;
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

        /// <summary>
        /// Bỏ qua bước intent-detection, sinh SQL trực tiếp cho câu hỏi.
        /// Dùng khi AI trả về need_sql=false với answer trống (câu hỏi quá phức tạp).
        /// </summary>
        private async Task<string> ForceSQLGenerationAsync(string question)
        {
            try
            {
                string prompt = $@"Bạn là SQL generator cho SQL Server. Sinh SELECT SQL cho yêu cầu dữ liệu dưới đây.

{DB_SCHEMA_SHORT}

Yêu cầu:
{question}

Quy tắc bắt buộc:
- Chỉ SELECT, không INSERT/UPDATE/DELETE
- Dùng đúng tên bảng/cột theo schema ở trên
- Is_Latest=1 cho MPR_Header, Is_Deleted=0 cho MPR_Details
- LEFT JOIN PO_Detail ON MPR_Detail_ID = Detail_ID (khi cần kiểm tra đã đặt hàng chưa)
- Nếu kiểm tra ""chưa đặt"": WHERE pd.MPR_Detail_ID IS NULL (sau LEFT JOIN)
- Nếu cột dimension = 0 thì dùng NULLIF(col, 0) hoặc CASE WHEN col > 0 THEN col END
- KHÔNG giới hạn TOP nếu yêu cầu đầy đủ, thêm ORDER BY phù hợp
- Chỉ trả về SQL thuần túy, KHÔNG markdown, KHÔNG giải thích

SQL:";

                string result = await CallRouterAsync(prompt, temperature: 0.05f);
                result = result.Trim();

                // Strip markdown fences nếu model bọc ```sql ... ```
                if (result.StartsWith("```"))
                {
                    int firstNewline = result.IndexOf('\n');
                    int lastFence    = result.LastIndexOf("```");
                    if (firstNewline >= 0 && lastFence > firstNewline)
                        result = result.Substring(firstNewline + 1, lastFence - firstNewline - 1).Trim();
                }

                // Chỉ chấp nhận nếu bắt đầu bằng SELECT hoặc WITH
                if (result.StartsWith("SELECT", StringComparison.OrdinalIgnoreCase) ||
                    result.StartsWith("WITH",   StringComparison.OrdinalIgnoreCase))
                    return result;

                return "";
            }
            catch { return ""; }
        }

        // ── Tự động sửa SQL khi gặp lỗi — tối đa 3 lần, tích lũy lỗi ──
        private async Task<string> FixSQLAsync(string badSql, string errorMsg, int attempt = 1)
        {
            try
            {
                string attemptNote = attempt > 1
                    ? $"\n⚠️ Đây là lần sửa thứ {attempt}. Các lỗi trước chưa giải quyết được — hãy xem xét kỹ hơn.\n"
                    : "";

                string fixPrompt = $@"SQL sau gặp lỗi khi chạy trên SQL Server:{attemptNote}

Lỗi:
{errorMsg}

SQL lỗi:
{badSql}

Schema chính xác (tên bảng và cột thực tế):
{DB_SCHEMA_SHORT}

Yêu cầu:
1. Sửa đúng tên bảng/cột theo schema ở trên.
2. Nếu bảng/cột không tồn tại trong schema → dùng bảng/cột thay thế phù hợp nhất.
3. Giữ nguyên logic truy vấn — chỉ sửa tên bảng/cột sai.
4. Chỉ trả về SQL thuần túy, không markdown, không giải thích.

SQL đã sửa:";

                string result = await CallRouterAsync(fixPrompt, temperature: 0.05f);
                // Loại bỏ markdown nếu model trả về có bọc ```sql ... ```
                result = result.Trim();
                if (result.StartsWith("```"))
                {
                    int firstNewline = result.IndexOf('\n');
                    int lastBacktick = result.LastIndexOf("```");
                    if (firstNewline >= 0 && lastBacktick > firstNewline)
                        result = result.Substring(firstNewline + 1, lastBacktick - firstNewline - 1).Trim();
                }
                return result;
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
            string historyCtx = BuildHistoryContext(6);

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
                    max_tokens = 4096,
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
                    max_tokens = 4096,
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
            if (dt.Rows.Count > maxRows)
                sb.AppendLine($"[QUAN TRỌNG: Truy vấn trả về {dt.Rows.Count} dòng — AI chỉ đọc được {maxRows} dòng đầu tiên. " +
                              $"Các con số tổng hợp (SUM, COUNT...) phải được tính trong SQL, không đếm tay từ danh sách này]");
            else
                sb.AppendLine($"Kết quả: {dt.Rows.Count} dòng / {dt.Columns.Count} cột");

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
                sb.AppendLine($"--- [HẾT {maxRows} dòng hiển thị / tổng {dt.Rows.Count} dòng] ---");

            return sb.ToString();
        }

        // ════════════════════════════════════════════════════════════════════
        // CHẾ ĐỘ CHAT TỰ DO — không truy vấn DB, AI trả lời như chatbot thường
        // ════════════════════════════════════════════════════════════════════

        public async Task<string> AskFreeAsync(string userQuestion, Action<string> onChunk = null)
        {
            string historyCtx = BuildHistoryContext(6);
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
                string truncated = text.Length > 500
                    ? text.Substring(0, 497) + "..."
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

                    // 4. Đợt thanh toán đã đến hạn nhưng chưa trả
                    var cmdPay = new SqlCommand(@"
                        SELECT COUNT(*) FROM PO_Payment_Schedule
                        WHERE Status NOT IN ('Paid','Cancelled')
                          AND Due_Date <= GETDATE()", conn);
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
                        ? $"💰 **{payPending} đợt thanh toán** đã đến hạn, chưa thanh toán"
                        : "✅ Không có đợt thanh toán nào đến hạn");

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

        // ════════════════════════════════════════════════════════════════════
        // AI SKILL — lưu skill (nút nhanh / lệnh tắt / workflow) vào bảng AI_Skill
        // ════════════════════════════════════════════════════════════════════

        public record SkillItem(int Id, string Name, string Type, string Slash,
                                string Template, string Description, string Icon, int Order,
                                string ParentSlash = "");

        private static List<SkillItem> _skillCache = null;
        private static DateTime _skillCacheTime = DateTime.MinValue;
        private static readonly TimeSpan SKILL_CACHE_TTL = TimeSpan.FromMinutes(5);
        private static void InvalidateSkillCache() => _skillCache = null;

        private static SqlConnection CreateSkillConn()
        {
            var b = new SqlConnectionStringBuilder(
                DatabaseHelper.GetConnection().ConnectionString)
            { MultipleActiveResultSets = true, ApplicationName = "MPR_AI_Skill" };
            return new SqlConnection(b.ConnectionString);
        }

        public static List<SkillItem> GetSkills(string type = null, bool forceRefresh = false)
        {
            if (!forceRefresh && _skillCache != null
                && DateTime.Now - _skillCacheTime < SKILL_CACHE_TTL)
            {
                return type == null ? _skillCache
                    : _skillCache.FindAll(s => s.Type == type);
            }
            var result = new List<SkillItem>();
            try
            {
                using var conn = CreateSkillConn();
                conn.Open();
                var dt = new DataTable();
                new SqlDataAdapter(new SqlCommand(
                    "SELECT Skill_ID, Skill_Name, Skill_Type, Slash_Command, " +
                    "Prompt_Template, Description, Icon, Sort_Order, Parent_Slash " +
                    "FROM AI_Skill WHERE Is_Active = 1 ORDER BY Sort_Order, Skill_ID", conn)).Fill(dt);
                foreach (DataRow r in dt.Rows)
                    result.Add(new SkillItem(
                        Convert.ToInt32(r["Skill_ID"]),
                        r["Skill_Name"].ToString(),
                        r["Skill_Type"].ToString(),
                        r["Slash_Command"] == DBNull.Value ? "" : r["Slash_Command"].ToString(),
                        r["Prompt_Template"].ToString(),
                        r["Description"] == DBNull.Value ? "" : r["Description"].ToString(),
                        r["Icon"].ToString(),
                        Convert.ToInt32(r["Sort_Order"]),
                        r["Parent_Slash"] == DBNull.Value ? "" : r["Parent_Slash"].ToString()));
                _skillCache = result;
                _skillCacheTime = DateTime.Now;
            }
            catch (Exception ex)
            { System.Diagnostics.Debug.WriteLine($"[AI_Skill] GetSkills: {ex.Message}"); }
            return type == null ? result : result.FindAll(s => s.Type == type);
        }

        public static SkillItem GetBySlashCommand(string slash)
        {
            if (string.IsNullOrWhiteSpace(slash)) return null;
            slash = slash.Trim();
            return GetSkills()?.Find(s =>
                s.Type == "slash" &&
                !string.IsNullOrEmpty(s.Slash) &&
                s.Slash.Equals(slash, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>Lấy danh sách sub-skill thuộc một lệnh cha.</summary>
        public static List<SkillItem> GetSubSkills(string parentSlash)
        {
            if (string.IsNullOrWhiteSpace(parentSlash)) return [];
            return GetSkills()?.FindAll(s =>
                !string.IsNullOrEmpty(s.ParentSlash) &&
                s.ParentSlash.Equals(parentSlash.Trim(), StringComparison.OrdinalIgnoreCase))
                ?? [];
        }

        public static string AddSkill(string name, string type, string template,
            string icon = "⚡", string desc = "", string slashCmd = "", int order = 0,
            string parentSlash = "")
        {
            name = name?.Trim() ?? ""; template = template?.Trim() ?? "";
            if (string.IsNullOrEmpty(name)) return "⚠️ Tên skill không được trống.";
            if (string.IsNullOrEmpty(template)) return "⚠️ Template không được trống.";
            try
            {
                using var conn = CreateSkillConn();
                conn.Open();
                var chk = new SqlCommand(
                    "SELECT COUNT(*) FROM AI_Skill WHERE Skill_Name = @n AND Is_Active = 1", conn);
                chk.Parameters.AddWithValue("@n", name);
                if (Convert.ToInt32(chk.ExecuteScalar()) > 0)
                    return $"⚠️ Skill \"{name}\" đã tồn tại.";
                var ins = new SqlCommand(
                    "INSERT INTO AI_Skill (Skill_Name, Skill_Type, Slash_Command, Prompt_Template, " +
                    "Description, Icon, Sort_Order, Parent_Slash) " +
                    "VALUES (@n, @t, @sl, @pt, @d, @ic, @so, @ps)", conn);
                ins.Parameters.AddWithValue("@n", name);
                ins.Parameters.AddWithValue("@t", type);
                ins.Parameters.AddWithValue("@sl", string.IsNullOrEmpty(slashCmd) ? (object)DBNull.Value : slashCmd);
                ins.Parameters.AddWithValue("@pt", template);
                ins.Parameters.AddWithValue("@d", string.IsNullOrEmpty(desc) ? (object)DBNull.Value : desc);
                ins.Parameters.AddWithValue("@ic", string.IsNullOrEmpty(icon) ? "⚡" : icon);
                ins.Parameters.AddWithValue("@so", order);
                ins.Parameters.AddWithValue("@ps", string.IsNullOrEmpty(parentSlash) ? (object)DBNull.Value : parentSlash);
                ins.ExecuteNonQuery();
                InvalidateSkillCache();
                return $"✅ Đã thêm skill: \"{name}\"";
            }
            catch (Exception ex) { return $"⚠️ Lỗi: {ex.Message}"; }
        }

        public static string RemoveSkill(int skillId)
        {
            try
            {
                using var conn = CreateSkillConn();
                conn.Open();
                var get = new SqlCommand(
                    "SELECT Skill_Name FROM AI_Skill WHERE Skill_ID = @id", conn);
                get.Parameters.AddWithValue("@id", skillId);
                string name = get.ExecuteScalar()?.ToString() ?? "";
                var del = new SqlCommand(
                    "UPDATE AI_Skill SET Is_Active = 0 WHERE Skill_ID = @id", conn);
                del.Parameters.AddWithValue("@id", skillId);
                del.ExecuteNonQuery();
                InvalidateSkillCache();
                return $"✅ Đã xóa skill: \"{name}\"";
            }
            catch (Exception ex) { return $"⚠️ Lỗi: {ex.Message}"; }
        }

        public static void EnsureSkillTableAndSeed()
        {
            try
            {
                using var conn = CreateSkillConn();
                conn.Open();

                // Tạo bảng nếu chưa có (bao gồm cột Parent_Slash)
                new SqlCommand(@"
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'AI_Skill')
BEGIN
    CREATE TABLE AI_Skill (
        Skill_ID        int IDENTITY(1,1) PRIMARY KEY,
        Skill_Name      nvarchar(100)  NOT NULL,
        Skill_Type      nvarchar(20)   NOT NULL,
        Slash_Command   nvarchar(50)   NULL,
        Prompt_Template nvarchar(max)  NOT NULL,
        Description     nvarchar(200)  NULL,
        Icon            nvarchar(10)   NOT NULL DEFAULT N'⚡',
        Sort_Order      int            NOT NULL DEFAULT 0,
        Is_Active       bit            NOT NULL DEFAULT 1,
        Parent_Slash    nvarchar(50)   NULL,
        Created_By      nvarchar(100)  NOT NULL DEFAULT 'User',
        Created_Date    datetime       NOT NULL DEFAULT GETDATE()
    )
END", conn).ExecuteNonQuery();

                // Thêm cột Parent_Slash nếu bảng cũ chưa có
                new SqlCommand(@"
IF NOT EXISTS (
    SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME = 'AI_Skill' AND COLUMN_NAME = 'Parent_Slash'
)
    ALTER TABLE AI_Skill ADD Parent_Slash nvarchar(50) NULL", conn).ExecuteNonQuery();

                // ── Patch: thêm {project_code} filter vào các skill MPR/PO (chạy mỗi startup) ──
                new SqlCommand(@"
UPDATE AI_Skill
SET Prompt_Template = Prompt_Template
    + N' Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.'
WHERE Is_Active = 1
  AND CHARINDEX(N'{project_code}', Prompt_Template) = 0
  AND Skill_Name IN (
    N'📋 MPR chưa có PO', N'🚨 PO quá hạn giao', N'💰 Thanh toán quá hạn',
    N'PO tháng này', N'PO quá hạn giao', N'PO chưa thanh toán',
    N'PO đang giao hàng', N'PO theo dự án',
    N'MPR chưa đặt hàng', N'MPR đặt hàng một phần',
    N'MPR revision mới nhất', N'MPR theo trạng thái',
    N'Thanh toán tháng này', N'Lịch thanh toán sắp tới',
    N'📈 KPI tháng này', N'📊 /sosanh', N'🔎 /vattu'
  )", conn).ExecuteNonQuery();

                // Seed chỉ khi bảng còn trống
                var count = Convert.ToInt32(
                    new SqlCommand("SELECT COUNT(*) FROM AI_Skill", conn).ExecuteScalar());
                if (count > 0) return;

                // (name, type, slash, template, desc, icon, order, parentSlash)
                var seeds = new[]
                {
                    // ── Quick ──────────────────────────────────────────────
                    ("📋 MPR chưa có PO",       "quick",    "",         "Danh sách MPR mới nhất (Is_Latest=1) chưa có PO nào được tạo. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",                       "MPR chưa được đặt hàng",     "📋", 1,  ""),
                    ("🚨 PO quá hạn giao",      "quick",    "",         "Danh sách PO chưa giao đủ và đã quá ngày Expected_Delivery, kèm tên nhà cung cấp và số ngày trễ. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",           "PO trễ hạn giao hàng",       "🚨", 2,  ""),
                    ("💰 Thanh toán quá hạn",   "quick",    "",         "Danh sách PO có lịch thanh toán (PO_Payment_Schedule) đã quá Due_Date nhưng Status chưa hoàn thành, kèm số tiền và tên NCC. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.", "Thanh toán chờ duyệt",       "💰", 3,  ""),
                    ("✅ RIR chờ QC",            "quick",    "",         "Danh sách RIR_head có Status chưa hoàn thành, kèm PONo và ngày tạo, sắp xếp mới nhất trước",                                                                    "RIR chờ kiểm tra",           "✅", 4,  ""),

                    // ── Slash cha ──────────────────────────────────────────
                    ("📦 /po",                  "slash",    "/po",      "Danh sách tất cả PO của dự án {param}, kèm PONo, tên nhà cung cấp, Total_Amount, Status, Expected_Delivery, sắp xếp theo PO_Date giảm dần. Nếu {param} trống hoặc {param}='All' thì lấy tất cả dự án.", "Danh sách PO theo dự án",    "📦", 10, ""),
                    ("🔍 /mpr",                 "slash",    "/mpr",     "Danh sách tất cả MPR (Is_Latest=1) của dự án {param}, kèm MPR_No, Status, Required_Date, số lượng item, sắp xếp theo Created_Date giảm dần. Nếu {param} trống hoặc {param}='All' thì lấy tất cả dự án.",          "Danh sách MPR theo dự án",   "🔍", 11, ""),
                    ("🏭 /ncc",                 "slash",    "/ncc",     "Danh sách nhà cung cấp đang hoạt động (IsActive=1) kèm số PO đang mở và tổng giá trị",                                                                           "Nhà cung cấp hoạt động",     "🏭", 12, ""),
                    ("📊 /kho",                 "slash",    "/kho",     "Tồn kho hiện tại từ vw_Warehouse_Stock_V2, nhóm theo vật tư và dự án",                                                                                           "Tồn kho hiện tại",           "📊", 13, ""),
                    ("💳 /tt",                  "slash",    "/tt",      "Tiến độ thanh toán từ vw_PO_Payment_Summary: % đã trả, còn lại, ngày đến hạn tiếp theo",                                                                        "Tiến độ thanh toán",         "💳", 14, ""),
                    ("🧾 /rir",                 "slash",    "/rir",     "Danh sách RIR tháng này kèm kết quả kiểm tra và PONo",                                                                                                           "RIR kiểm tra nhập hàng",     "🧾", 15, ""),

                    // ── Sub-skill của /po ─────────────────────────────────
                    ("PO tháng này",            "slash",    "",         "Danh sách PO được tạo trong tháng này, kèm tên NCC, tổng giá trị và trạng thái. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",                            "PO trong tháng hiện tại",    "📅", 20, "/po"),
                    ("PO quá hạn giao",         "slash",    "",         "Danh sách PO đã quá Expected_Delivery nhưng chưa giao đủ hàng (Received_Qty < Qty_Per_Sheet), kèm số ngày trễ và tên NCC. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.", "PO trễ hạn giao hàng",       "🚨", 21, "/po"),
                    ("PO chưa thanh toán",      "slash",    "",         "Danh sách PO chưa thanh toán đủ (Amount_Remaining > 0 từ vw_PO_Payment_Summary), kèm số tiền còn lại. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",      "PO còn nợ thanh toán",       "💸", 22, "/po"),
                    ("PO đang giao hàng",       "slash",    "",         "Danh sách PO có hàng đang giao (có trong PO_DeliveryTracking với Status chưa hoàn thành). Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",                  "PO đang trong quá trình giao","🚚", 23, "/po"),
                    ("PO theo dự án",           "slash",    "",         "Thống kê số PO và tổng giá trị theo từng dự án (Project_Name) trong PO_head. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",                              "PO nhóm theo dự án",         "📁", 24, "/po"),

                    // ── Sub-skill của /mpr ────────────────────────────────
                    ("MPR chưa đặt hàng",       "slash",    "",         "Danh sách MPR (Is_Latest=1) chưa có PO nào liên kết, kèm tên dự án và ngày cần hàng. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",                      "MPR chưa có PO",             "❌", 30, "/mpr"),
                    ("MPR đặt hàng một phần",   "slash",    "",         "Danh sách MPR có một số item đã có PO nhưng chưa đặt hết, kèm % item đã đặt. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",                              "MPR đặt hàng chưa đủ",       "⚠️", 31, "/mpr"),
                    ("MPR revision mới nhất",   "slash",    "",         "Danh sách MPR có nhiều revision, chỉ lấy bản mới nhất (Is_Latest=1), kèm số Rev và ngày tạo. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",               "MPR bản revision cuối",      "🔄", 32, "/mpr"),
                    ("MPR theo trạng thái",     "slash",    "",         "Thống kê số lượng MPR theo từng trạng thái (Status), kèm danh sách chi tiết. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",                              "MPR nhóm theo trạng thái",   "📊", 33, "/mpr"),

                    // ── Sub-skill của /ncc ────────────────────────────────
                    ("NCC có công nợ",          "slash",    "",         "Danh sách NCC có công nợ chưa thanh toán (Total_Debt > 0 từ vw_Supplier_Debt_Summary), sắp xếp theo nợ cao nhất",                                               "NCC còn nợ",                 "💰", 40, "/ncc"),
                    ("NCC theo loại",           "slash",    "",         "Thống kê nhà cung cấp nhóm theo Supplier_Type, kèm số PO và tổng giá trị mỗi loại",                                                                             "NCC phân loại",              "🏷️", 41, "/ncc"),
                    ("NCC chứng chỉ",           "slash",    "",         "Danh sách NCC kèm cột Cert trong bảng Suppliers (KHÔNG dùng Supplier_Certificates — không tồn tại). Chỉ lấy NCC có Cert không rỗng, hiển thị Company_Name, Short_Name, Cert, Contact_Person, Contact_Phone",                      "Chứng chỉ nhà cung cấp",     "📜", 42, "/ncc"),
                    ("NCC tài khoản ngân hàng", "slash",    "",         "Danh sách NCC kèm thông tin tài khoản ngân hàng trực tiếp từ bảng Suppliers: Bank_Account, Bank_Name (KHÔNG dùng Supplier_Bank_Accounts — không tồn tại). Chỉ lấy NCC có Bank_Account không rỗng, kèm Company_Name, Tax_Code, Contact_Person",                      "Tài khoản thanh toán NCC",   "🏦", 43, "/ncc"),

                    // ── Sub-skill của /kho ────────────────────────────────
                    ("Tồn kho hiện tại",        "slash",    "",         "Tồn kho hiện tại từ vw_Warehouse_Stock_V2 (Qty_Stock > 0), nhóm theo vật tư và vị trí lưu kho",                                                                 "Vật tư đang tồn kho",        "📦", 50, "/kho"),
                    ("Vào kho tháng này",       "slash",    "",         "Danh sách vật tư nhập kho (Warehouse_Import) trong tháng này, kèm PO nguồn và dự án",                                                                           "Nhập kho tháng hiện tại",    "📥", 51, "/kho"),
                    ("Xuất kho tháng này",      "slash",    "",         "Danh sách vật tư xuất kho (Warehouse_Export) trong tháng này, kèm mục đích xuất và dự án",                                                                      "Xuất kho tháng hiện tại",    "📤", 52, "/kho"),
                    ("Chuyển kho",              "slash",    "",         "Danh sách giao dịch chuyển kho (ProjectMaterialTransformTransaction) trong tháng này, kèm vị trí cũ, vị trí mới và số lượng",                                   "Lịch sử chuyển kho",         "🔄", 53, "/kho"),

                    // ── Sub-skill của /tt (thanh toán) ────────────────────
                    ("Thanh toán tháng này",    "slash",    "",         "Danh sách các lần thanh toán (PO_Payment_History) đã thực hiện trong tháng này, kèm số tiền, ngân hàng và số giao dịch. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.", "Thanh toán trong tháng",     "💳", 60, "/tt"),
                    ("Lịch thanh toán sắp tới", "slash",    "",         "Danh sách kế hoạch thanh toán (PO_Payment_Schedule) chưa hoàn thành, sắp đến hạn trong 30 ngày tới, kèm PO và số tiền. Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.", "Kế hoạch TT sắp đến hạn",   "📅", 61, "/tt"),
                    ("Công nợ theo NCC",        "slash",    "",         "Tổng hợp công nợ nhà cung cấp từ vw_Supplier_Debt_Summary, kèm tổng nợ, số PO và ngày thanh toán gần nhất",                                                     "Công nợ NCC",                "🧾", 62, "/tt"),

                    // ── Workflow ───────────────────────────────────────────
                    ("🔬 Phân tích NCC",        "workflow", "",         "Phân tích toàn diện nhà cung cấp {param}: danh sách PO, tổng giá trị, tiến độ giao hàng, tình trạng thanh toán và công nợ. Nếu {param} trống hãy hỏi tên NCC.", "Phân tích tổng hợp NCC",     "🔬", 70, ""),
                    ("📁 Theo dõi dự án",       "workflow", "",         "Toàn bộ thông tin dự án {param}: danh sách MPR, PO liên quan, tổng giá trị đặt hàng, tiến độ nhập kho. Nếu {param} trống hãy hỏi mã hoặc tên dự án.",           "Tổng hợp theo dự án",        "📁", 71, ""),
                    ("📦 Theo dõi PO",          "workflow", "",         "Theo dõi chi tiết PO {param}: từng dòng vật tư, tiến độ giao hàng, lịch thanh toán và nhập kho. Nếu {param} trống hãy hỏi số PO.",                              "Chi tiết một PO",            "📦", 72, ""),

                    // ── Quick: KPI tổng quan ───────────────────────────────
                    ("📈 KPI tháng này",        "quick",    "",
                        "Tổng hợp KPI tháng hiện tại: (1) số PO mới tạo và tổng giá trị, (2) số MPR mới, " +
                        "(3) số RIR hoàn thành, (4) tổng tiền đã thanh toán (PO_Payment_History tháng này), " +
                        "(5) số PO quá hạn giao. Mỗi KPI là 1 dòng với số liệu cụ thể. " +
                        "Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",
                        "KPI tổng quan tháng này",    "📈", 5,  ""),

                    // ── Quick: tồn kho dưới ngưỡng ────────────────────────
                    ("⚠️ Tồn kho thấp",         "quick",    "",
                        "Danh sách vật tư từ vw_Warehouse_Stock_V2 có Qty_Stock > 0 nhưng nhỏ hơn 50 " +
                        "(hoặc dưới ngưỡng bình thường), kèm vị trí kho và dự án. Sắp xếp theo Qty_Stock tăng dần.",
                        "Cảnh báo tồn kho thấp",      "⚠️", 6,  ""),

                    // ── Slash: so sánh tháng này vs tháng trước ───────────
                    ("📊 /sosanh",              "slash",    "/sosanh",
                        "So sánh tháng này vs tháng trước: số PO mới, tổng giá trị PO, " +
                        "số RIR, số MPR, tổng tiền thanh toán. " +
                        "Dùng MONTH/YEAR(GETDATE()) và MONTH/YEAR(DATEADD(MONTH,-1,GETDATE())) để lọc. " +
                        "Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",
                        "So sánh tháng này vs tháng trước","📊", 16, ""),

                    // ── Slash: tìm kiếm vật tư ────────────────────────────
                    ("🔎 /vattu",               "slash",    "/vattu",
                        "Tìm kiếm vật tư trong hệ thống: tìm trong MPR_Details (item_name, Material, Is_Deleted=0) " +
                        "VÀ Warehouse_Import (Item_Name, Material) VÀ vw_Warehouse_Stock_V2 (Item_Name). " +
                        "Tổng hợp: đã đặt bao nhiêu, đã nhập kho bao nhiêu, còn tồn bao nhiêu. " +
                        "Lọc dự án: {project_code}. Nếu {project_code} trống thì lấy tất cả dự án.",
                        "Tìm kiếm và tổng hợp vật tư",  "🔎", 17, ""),

                    // ── Workflow: tiến độ dự án ────────────────────────────
                    ("🏗️ Tiến độ dự án",        "workflow", "",
                        "Báo cáo tiến độ đầy đủ cho dự án {param}: " +
                        "Tổng MPR → đã có PO / chưa có PO, " +
                        "PO theo trạng thái (Draft/Approved/Sent/Partial/Completed), " +
                        "% item đã nhập kho vs tổng đặt, " +
                        "đợt thanh toán còn nợ. " +
                        "Nếu {param} trống hãy hỏi mã dự án (Project_Code).",
                        "Tiến độ tổng thể dự án",     "🏗️", 73, ""),
                };

                foreach (var (name, type, slash, template, desc, icon, order, parentSlash) in seeds)
                {
                    var ins = new SqlCommand(
                        "INSERT INTO AI_Skill (Skill_Name, Skill_Type, Slash_Command, Prompt_Template, Description, Icon, Sort_Order, Parent_Slash) " +
                        "VALUES (@n, @t, @sl, @pt, @d, @ic, @so, @ps)", conn);
                    ins.Parameters.AddWithValue("@n", name);
                    ins.Parameters.AddWithValue("@t", type);
                    ins.Parameters.AddWithValue("@sl", string.IsNullOrEmpty(slash) ? (object)DBNull.Value : slash);
                    ins.Parameters.AddWithValue("@pt", template);
                    ins.Parameters.AddWithValue("@d", (object)desc ?? DBNull.Value);
                    ins.Parameters.AddWithValue("@ic", icon);
                    ins.Parameters.AddWithValue("@so", order);
                    ins.Parameters.AddWithValue("@ps", string.IsNullOrEmpty(parentSlash) ? (object)DBNull.Value : parentSlash);
                    ins.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            { System.Diagnostics.Debug.WriteLine($"[AI_Skill] EnsureTableAndSeed: {ex.Message}"); }
        }

        // ════════════════════════════════════════════════════════════════════
        // DANH SÁCH DỰ ÁN — dùng cho project picker dialog
        // ════════════════════════════════════════════════════════════════════

        // ── Cache project codes — tránh query DB mỗi lần gõ phím ──────
        private static List<string> _projectCodesCache = null;
        private static DateTime _projectCodesCacheTime = DateTime.MinValue;
        private static readonly TimeSpan PROJECT_CODES_CACHE_TTL = TimeSpan.FromMinutes(3);

        /// <summary>
        /// Lấy danh sách mã dự án từ DB (MPR_Header + PO_head), bỏ trùng, sắp xếp A-Z.
        /// Kết quả được cache 3 phút để popup không bị lag khi gõ phím.
        /// </summary>
        public static List<string> GetProjectCodes(bool forceRefresh = false)
        {
            if (!forceRefresh && _projectCodesCache != null
                && DateTime.Now - _projectCodesCacheTime < PROJECT_CODES_CACHE_TTL)
                return _projectCodesCache;

            var result = new List<string>();
            try
            {
                var builder = new SqlConnectionStringBuilder(
                    DatabaseHelper.GetConnection().ConnectionString)
                { MultipleActiveResultSets = true, ApplicationName = "MPR_AI_ProjectCodes" };
                using var conn = new SqlConnection(builder.ConnectionString);
                conn.Open();
                var dt = new DataTable();
                new SqlDataAdapter(new SqlCommand(@"
                    SELECT DISTINCT LTRIM(RTRIM(ProjectCode)) AS Code
                    FROM PO_head
                    WHERE ProjectCode IS NOT NULL AND LTRIM(RTRIM(ProjectCode)) <> ''
                    UNION
                    SELECT DISTINCT LTRIM(RTRIM(Project_Code))
                    FROM MPR_Header
                    WHERE Project_Code IS NOT NULL AND LTRIM(RTRIM(Project_Code)) <> ''
                    ORDER BY Code", conn)).Fill(dt);
                foreach (DataRow r in dt.Rows)
                {
                    string v = r[0]?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(v)) result.Add(v);
                }
                _projectCodesCache = result;
                _projectCodesCacheTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GetProjectCodes] {ex.Message}");
            }
            return result;
        }

        // ════════════════════════════════════════════════════════════════════
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