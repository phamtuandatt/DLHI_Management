using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace MPR_Managerment.Services
{
    public class ZaloImportRow
    {
        public int RowIndex { get; set; }
        public string PONo { get; set; }
        public string EcountNo { get; set; }
        public string TitleEN { get; set; }
        public string GWNo { get; set; }
        public decimal? Amount { get; set; }
        public decimal? VAT { get; set; }
        public decimal? FinalAmount { get; set; }
        public decimal? Dot1 { get; set; }
        public decimal? Dot2 { get; set; }
        public decimal? Dot3 { get; set; }
        public decimal? Dot4 { get; set; }
        public decimal? Dot5 { get; set; }
        public decimal? Dot6 { get; set; }
        public decimal? Dot7 { get; set; }
        public decimal? Dot8 { get; set; }
        public string FTCash { get; set; }
        public DateTime? PaidDate { get; set; }
        public string ProgressStatus { get; set; }
        public string Note { get; set; }
    }

    public class ZaloImportResult
    {
        public DateTime FileDate { get; set; }
        public string FilePath { get; set; }
        public int TotalRows { get; set; }
        public int InsertedRows { get; set; }
        public int UpdatedRows { get; set; }
        public string Error { get; set; }
        public bool Success => string.IsNullOrEmpty(Error);
    }

    public class ZaloImportService
    {
        // Folder Zalo auto-download
        public static readonly string ZaloDownloadFolder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                         "Temp", "Zalo Temp", "TempDownloads");

        private const string FILE_PREFIX = "0. 품의서 ";
        private const string SHEET_NAME = "Payment list";

        // Trả về tất cả file 품의서 trong folder, sắp xếp mới nhất trước
        public static List<(string path, DateTime date)> FindImportFiles(string folder = null)
        {
            folder ??= ZaloDownloadFolder;
            var result = new List<(string, DateTime)>();
            if (!Directory.Exists(folder)) return result;

            foreach (var f in Directory.GetFiles(folder, "0. 품의서 *.xlsx"))
            {
                string fname = Path.GetFileNameWithoutExtension(f);
                string datePart = fname.Substring(FILE_PREFIX.Length).Trim();
                if (DateTime.TryParseExact(datePart, "yyyy-MM-dd",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out DateTime dt))
                {
                    result.Add((f, dt));
                }
            }
            result.Sort((a, b) => b.Item2.CompareTo(a.Item2));
            return result;
        }

        // Đọc dữ liệu từ file Excel — chỉ sheet "Payment list", từ row 4
        public static List<ZaloImportRow> ReadFile(string filePath)
        {
            var rows = new List<ZaloImportRow>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var pkg = new ExcelPackage(new FileInfo(filePath));
            var ws = pkg.Workbook.Worksheets[SHEET_NAME];
            if (ws == null)
                throw new Exception($"Không tìm thấy sheet '{SHEET_NAME}' trong file.");

            int lastRow = ws.Dimension?.End.Row ?? 0;

            for (int r = 4; r <= lastRow; r++)
            {
                // Bỏ qua dòng hoàn toàn trống (kiểm tra H hoặc J có giá trị)
                var h = GetDecimal(ws.Cells[r, 8]);
                var j = GetDecimal(ws.Cells[r, 10]);
                string poNo = CleanStr(ws.Cells[r, 3].Text);
                string gwNo = CleanStr(ws.Cells[r, 7].Text);

                if (h == null && j == null && string.IsNullOrEmpty(poNo) && string.IsNullOrEmpty(gwNo))
                    continue;

                rows.Add(new ZaloImportRow
                {
                    RowIndex     = r,
                    PONo         = poNo,
                    EcountNo     = CleanStr(ws.Cells[r, 4].Text),
                    TitleEN      = CleanStr(ws.Cells[r, 5].Text),
                    GWNo         = CleanStr(ws.Cells[r, 7].Text),
                    Amount       = h,
                    VAT          = GetDecimal(ws.Cells[r, 9]),
                    FinalAmount  = j,
                    Dot1         = GetDecimal(ws.Cells[r, 12]),
                    Dot2         = GetDecimal(ws.Cells[r, 13]),
                    Dot3         = GetDecimal(ws.Cells[r, 14]),
                    Dot4         = GetDecimal(ws.Cells[r, 15]),
                    Dot5         = GetDecimal(ws.Cells[r, 16]),
                    Dot6         = GetDecimal(ws.Cells[r, 17]),
                    Dot7         = GetDecimal(ws.Cells[r, 18]),
                    Dot8         = GetDecimal(ws.Cells[r, 19]),
                    FTCash       = ws.Cells[r, 22].Text?.Trim(),
                    PaidDate  = GetDate(ws.Cells[r, 23]),
                    ProgressStatus = ws.Cells[r, 27].Text?.Trim(),
                });
            }
            return rows;
        }

        // Import vào DB, ghi đè nếu trùng (File_Date + Row_Index)
        public static ZaloImportResult ImportToDB(string filePath, DateTime fileDate)
        {
            var result = new ZaloImportResult { FilePath = filePath, FileDate = fileDate };
            try
            {
                var rows = ReadFile(filePath);
                result.TotalRows = rows.Count;

                EnsureTable();

                using var conn = DatabaseHelper.GetConnection();
                conn.Open();

                foreach (var row in rows)
                {
                    // Kiểm tra đã tồn tại chưa
                    var chk = new SqlCommand(
                        "SELECT COUNT(1) FROM Zalo_PaymentImport WHERE File_Date=@fd AND Row_Index=@ri", conn);
                    chk.Parameters.AddWithValue("@fd", fileDate.Date);
                    chk.Parameters.AddWithValue("@ri", row.RowIndex);
                    bool exists = (int)chk.ExecuteScalar() > 0;

                    if (exists)
                    {
                        var upd = new SqlCommand(@"
                            UPDATE Zalo_PaymentImport SET
                                PO_No=@po, Ecount_No=@ec, Title_EN=@ti, GW_No=@gw,
                                Amount=@am, VAT=@vat, Final_Amount=@fa,
                                Dot1=@d1,Dot2=@d2,Dot3=@d3,Dot4=@d4,
                                Dot5=@d5,Dot6=@d6,Dot7=@d7,Dot8=@d8,
                                FT_Cash=@ftc, Paid_Date=@pd,
                                Progress_Status=@ps, Imported_At=GETDATE(), Source_File=@sf
                            WHERE File_Date=@fd AND Row_Index=@ri", conn);
                        FillParams(upd, row, fileDate, filePath);
                        upd.ExecuteNonQuery();
                        result.UpdatedRows++;
                    }
                    else
                    {
                        var ins = new SqlCommand(@"
                            INSERT INTO Zalo_PaymentImport
                                (File_Date,Row_Index,PO_No,Ecount_No,Title_EN,GW_No,
                                 Amount,VAT,Final_Amount,
                                 Dot1,Dot2,Dot3,Dot4,Dot5,Dot6,Dot7,Dot8,
                                 FT_Cash,Paid_Date,
                                 Progress_Status,Source_File)
                            VALUES
                                (@fd,@ri,@po,@ec,@ti,@gw,
                                 @am,@vat,@fa,
                                 @d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,
                                 @ftc,@pd,
                                 @ps,@sf)", conn);
                        FillParams(ins, row, fileDate, filePath);
                        ins.ExecuteNonQuery();
                        result.InsertedRows++;
                    }
                }
            }
            catch (Exception ex)
            {
                result.Error = ex.Message;
            }
            return result;
        }

        private static void FillParams(SqlCommand cmd, ZaloImportRow row, DateTime fileDate, string filePath)
        {
            cmd.Parameters.AddWithValue("@fd",  fileDate.Date);
            cmd.Parameters.AddWithValue("@ri",  row.RowIndex);
            // Loại bỏ TOÀN BỘ whitespace khỏi PO_No trước khi lưu vào DB
            string cleanPo = row.PONo == null ? null
                : new string(row.PONo.Where(c => !char.IsWhiteSpace(c)).ToArray());
            cmd.Parameters.AddWithValue("@po",  (object)(string.IsNullOrEmpty(cleanPo) ? null : cleanPo) ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ec",  (object)row.EcountNo ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ti",  (object)row.TitleEN ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@gw",  (object)row.GWNo  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@am",  (object)row.Amount  ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@vat", (object)row.VAT     ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@fa",  (object)row.FinalAmount ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@d1",  (object)row.Dot1 ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@d2",  (object)row.Dot2 ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@d3",  (object)row.Dot3 ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@d4",  (object)row.Dot4 ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@d5",  (object)row.Dot5 ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@d6",  (object)row.Dot6 ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@d7",  (object)row.Dot7 ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@d8",  (object)row.Dot8 ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ftc", (object)row.FTCash ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@pd",  (object)row.PaidDate ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@ps",  (object)row.ProgressStatus ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@sf",  (object)filePath ?? DBNull.Value);
        }

        public static void EnsureTable()
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            new SqlCommand(@"
                IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='Zalo_PaymentImport')
                CREATE TABLE Zalo_PaymentImport (
                    Import_ID       INT IDENTITY(1,1) PRIMARY KEY,
                    File_Date       DATE NOT NULL,
                    Row_Index       INT NOT NULL,
                    PO_No           NVARCHAR(200) NULL,
                    Ecount_No       NVARCHAR(100) NULL,
                    Title_EN        NVARCHAR(500) NULL,
                    GW_No           NVARCHAR(200) NULL,
                    Amount          DECIMAL(18,2) NULL,
                    VAT             DECIMAL(18,2) NULL,
                    Final_Amount    DECIMAL(18,2) NULL,
                    Dot1            DECIMAL(18,2) NULL,
                    Dot2            DECIMAL(18,2) NULL,
                    Dot3            DECIMAL(18,2) NULL,
                    Dot4            DECIMAL(18,2) NULL,
                    Dot5            DECIMAL(18,2) NULL,
                    Dot6            DECIMAL(18,2) NULL,
                    Dot7            DECIMAL(18,2) NULL,
                    Dot8            DECIMAL(18,2) NULL,
                    FT_Cash         NVARCHAR(500) NULL,
                    Paid_Date    DATE NULL,
                    Progress_Status NVARCHAR(100) NULL,
                    Imported_At     DATETIME DEFAULT GETDATE(),
                    Source_File     NVARCHAR(500) NULL,
                    CONSTRAINT UQ_ZaloImport UNIQUE (File_Date, Row_Index)
                );
                IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Zalo_PaymentImport' AND COLUMN_NAME='FT_Cash')
                    ALTER TABLE Zalo_PaymentImport ADD FT_Cash NVARCHAR(500) NULL;
                IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Zalo_PaymentImport' AND COLUMN_NAME='FT_Cash' AND CHARACTER_MAXIMUM_LENGTH < 500)
                    ALTER TABLE Zalo_PaymentImport ALTER COLUMN FT_Cash NVARCHAR(500) NULL;
                IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Zalo_PaymentImport' AND COLUMN_NAME='Paid_Date')
                    ALTER TABLE Zalo_PaymentImport ADD Paid_Date DATE NULL;
                IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Zalo_PaymentImport' AND COLUMN_NAME='Note')
                    ALTER TABLE Zalo_PaymentImport ADD Note NVARCHAR(500) NULL;
                UPDATE Zalo_PaymentImport SET PO_No = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(PO_No)), CHAR(160), ''), CHAR(9), ''), CHAR(13), ''), CHAR(10), ''), ' ', '');
                UPDATE Zalo_PaymentImport SET PO_No = NULL WHERE ISNULL(PO_No,'') = '';", conn).ExecuteNonQuery();
        }

        public static void SaveNote(int importId, string note)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand("UPDATE Zalo_PaymentImport SET Note=@n WHERE Import_ID=@id", conn);
            cmd.Parameters.AddWithValue("@n", (object)note ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@id", importId);
            cmd.ExecuteNonQuery();
        }

        private static DateTime? GetDate(ExcelRange cell)
        {
            if (cell?.Value == null) return null;
            if (cell.Value is DateTime dt) return dt.Date;
            string s = cell.Text?.Trim();
            if (string.IsNullOrWhiteSpace(s)) return null;
            string[] fmts = { "yyyy-MM-dd", "dd/MM/yyyy", "MM/dd/yyyy", "yyyy/MM/dd" };
            if (DateTime.TryParseExact(s, fmts, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out DateTime d)) return d.Date;
            if (DateTime.TryParse(s, out DateTime d2)) return d2.Date;
            return null;
        }

        private static string CleanStr(string s)
        {
            if (s == null) return null;
            // Loại bỏ tất cả loại khoảng trắng: space, tab, non-breaking space (U+00A0), ZWSP...
            s = s.Replace(" ", " ").Replace("\t", " ").Trim();
            return string.IsNullOrEmpty(s) ? null : s;
        }

        private static decimal? GetDecimal(ExcelRange cell)
        {
            if (cell?.Value == null) return null;
            string s = cell.Text?.Replace(",", "").Trim();
            if (string.IsNullOrWhiteSpace(s)) return null;
            return decimal.TryParse(s, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out decimal d) ? d : null;
        }
    }
}
