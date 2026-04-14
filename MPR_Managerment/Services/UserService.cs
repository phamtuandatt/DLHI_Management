// ============================================================
//  FILE: Services/UserService.cs
//  Toàn bộ logic User + Permission
// ============================================================

using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;

namespace MPR_Managerment.Services
{
    public class UserService
    {
        // =====================================================
        //  HASH PASSWORD — SHA256
        // =====================================================
        public static string HashPassword(string password)
        {
            using var sha = SHA256.Create();
            byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(password));
            var sb = new StringBuilder();
            foreach (byte b in bytes) sb.Append(b.ToString("x2"));
            return sb.ToString();
        }

        // =====================================================
        //  LOGIN
        // =====================================================
        public LoginResult Login(string username, string password)
        {
            string hash = HashPassword(password);
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand("sp_Login", conn)
            {
                CommandType = System.Data.CommandType.StoredProcedure
            };
            cmd.Parameters.AddWithValue("@Username", username);
            cmd.Parameters.AddWithValue("@Password_Hash", hash);

            using var r = cmd.ExecuteReader();
            if (!r.Read())
                return new LoginResult { Success = false, Message = "Không có phản hồi từ server" };

            string result = r["Result"]?.ToString() ?? "";
            if (result != "SUCCESS")
            {
                string msg = result switch
                {
                    "INVALID_CREDENTIALS" => "Sai tên đăng nhập hoặc mật khẩu!",
                    "ACCOUNT_DISABLED" => "Tài khoản đã bị vô hiệu hóa!",
                    _ => "Đăng nhập thất bại!"
                };
                return new LoginResult { Success = false, Message = msg };
            }

            var user = new AppUser
            {
                User_ID = Convert.ToInt32(r["User_ID"]),
                Full_Name = r["Full_Name"]?.ToString() ?? "",
                Role_ID = Convert.ToInt32(r["Role_ID"]),
                Role_Name = r["Role_Name"]?.ToString() ?? "",
                Username = username,
                Must_Change_Password = Convert.ToBoolean(r["Must_Change_Password"])
            };
            r.Close();

            // Load quyền
            var perms = GetEffectivePermissions(user.User_ID, conn);

            return new LoginResult { Success = true, User = user, Permissions = perms };
        }

        // =====================================================
        //  LẤY QUYỀN HIỆU LỰC (từ view)
        // =====================================================
        public List<UserPermission> GetEffectivePermissions(int userId, SqlConnection? existingConn = null)
        {
            var list = new List<UserPermission>();
            bool ownsConn = existingConn == null;
            var conn = existingConn ?? DatabaseHelper.GetConnection();
            try
            {
                if (ownsConn) conn.Open();
                var cmd = new SqlCommand(
                    "SELECT * FROM vw_User_Effective_Permissions WHERE User_ID = @uid", conn);
                cmd.Parameters.AddWithValue("@uid", userId);
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    list.Add(new UserPermission
                    {
                        Module_ID = Convert.ToInt32(r["Module_ID"]),
                        Module_Code = r["Module_Code"]?.ToString() ?? "",
                        Module_Name = r["Module_Name"]?.ToString() ?? "",
                        Can_View = Convert.ToBoolean(r["Can_View"]),
                        Can_Create = Convert.ToBoolean(r["Can_Create"]),
                        Can_Edit = Convert.ToBoolean(r["Can_Edit"]),
                        Can_Delete = Convert.ToBoolean(r["Can_Delete"]),
                        Can_Export = Convert.ToBoolean(r["Can_Export"]),
                        Is_Custom_Override = Convert.ToBoolean(r["Is_Custom_Override"])
                    });
                }
            }
            finally { if (ownsConn) conn.Dispose(); }
            return list;
        }

        // =====================================================
        //  GET ALL USERS
        // =====================================================
        public List<AppUser> GetAll()
        {
            var list = new List<AppUser>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                SELECT u.*, r.Role_Name
                FROM Users u INNER JOIN Roles r ON r.Role_ID = u.Role_ID
                ORDER BY u.Username", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read()) list.Add(MapUser(r));
            return list;
        }

        // =====================================================
        //  GET ALL ROLES
        // =====================================================
        public List<Role> GetRoles()
        {
            var list = new List<Role>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand("SELECT * FROM Roles WHERE Is_Active=1 ORDER BY Role_ID", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new Role
                {
                    Role_ID = Convert.ToInt32(r["Role_ID"]),
                    Role_Name = r["Role_Name"]?.ToString() ?? "",
                    Description = r["Description"]?.ToString() ?? "",
                    Is_Active = Convert.ToBoolean(r["Is_Active"])
                });
            return list;
        }

        // =====================================================
        //  GET ALL MODULES
        // =====================================================
        public List<AppModule> GetModules()
        {
            var list = new List<AppModule>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand("SELECT * FROM Modules ORDER BY Sort_Order", conn);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                list.Add(new AppModule
                {
                    Module_ID = Convert.ToInt32(r["Module_ID"]),
                    Module_Code = r["Module_Code"]?.ToString() ?? "",
                    Module_Name = r["Module_Name"]?.ToString() ?? "",
                    Sort_Order = Convert.ToInt32(r["Sort_Order"])
                });
            return list;
        }

        // =====================================================
        //  INSERT USER
        // =====================================================
        public int InsertUser(AppUser user, string plainPassword, string createdBy)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                INSERT INTO Users (Username, Password_Hash, Full_Name, Email, Phone,
                                   Department, Role_ID, Is_Active, Must_Change_Password, Created_By)
                VALUES (@username, @hash, @fullname, @email, @phone,
                        @dept, @roleId, @active, @mustChange, @createdBy);
                SELECT SCOPE_IDENTITY();", conn);

            cmd.Parameters.AddWithValue("@username", user.Username);
            cmd.Parameters.AddWithValue("@hash", HashPassword(plainPassword));
            cmd.Parameters.AddWithValue("@fullname", user.Full_Name);
            cmd.Parameters.AddWithValue("@email", user.Email ?? "");
            cmd.Parameters.AddWithValue("@phone", user.Phone ?? "");
            cmd.Parameters.AddWithValue("@dept", user.Department ?? "");
            cmd.Parameters.AddWithValue("@roleId", user.Role_ID);
            cmd.Parameters.AddWithValue("@active", user.Is_Active);
            cmd.Parameters.AddWithValue("@mustChange", user.Must_Change_Password);
            cmd.Parameters.AddWithValue("@createdBy", createdBy);

            return Convert.ToInt32(cmd.ExecuteScalar());
        }

        // =====================================================
        //  UPDATE USER
        // =====================================================
        public void UpdateUser(AppUser user, string modifiedBy)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                UPDATE Users SET
                    Full_Name     = @fullname,
                    Email         = @email,
                    Phone         = @phone,
                    Department    = @dept,
                    Role_ID       = @roleId,
                    Is_Active     = @active,
                    Modified_Date = GETDATE(),
                    Modified_By   = @modBy
                WHERE User_ID = @userId", conn);

            cmd.Parameters.AddWithValue("@userId", user.User_ID);
            cmd.Parameters.AddWithValue("@fullname", user.Full_Name);
            cmd.Parameters.AddWithValue("@email", user.Email ?? "");
            cmd.Parameters.AddWithValue("@phone", user.Phone ?? "");
            cmd.Parameters.AddWithValue("@dept", user.Department ?? "");
            cmd.Parameters.AddWithValue("@roleId", user.Role_ID);
            cmd.Parameters.AddWithValue("@active", user.Is_Active);
            cmd.Parameters.AddWithValue("@modBy", modifiedBy);
            cmd.ExecuteNonQuery();
        }

        // =====================================================
        //  RESET PASSWORD (Admin)
        // =====================================================
        public void ResetPassword(int userId, string newPassword, string modifiedBy)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                UPDATE Users SET
                    Password_Hash        = @hash,
                    Must_Change_Password = 1,
                    Modified_Date        = GETDATE(),
                    Modified_By          = @modBy
                WHERE User_ID = @userId", conn);
            cmd.Parameters.AddWithValue("@hash", HashPassword(newPassword));
            cmd.Parameters.AddWithValue("@modBy", modifiedBy);
            cmd.Parameters.AddWithValue("@userId", userId);
            cmd.ExecuteNonQuery();
        }

        // =====================================================
        //  CHANGE PASSWORD (User tự đổi)
        // =====================================================
        public (bool success, string message) ChangePassword(int userId, string oldPwd, string newPwd)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand("sp_ChangePassword", conn)
            {
                CommandType = System.Data.CommandType.StoredProcedure
            };
            cmd.Parameters.AddWithValue("@User_ID", userId);
            cmd.Parameters.AddWithValue("@Old_Hash", HashPassword(oldPwd));
            cmd.Parameters.AddWithValue("@New_Hash", HashPassword(newPwd));

            using var r = cmd.ExecuteReader();
            if (r.Read())
                return (Convert.ToBoolean(r["Success"]), r["Message"]?.ToString() ?? "");
            return (false, "Không có phản hồi");
        }

        // =====================================================
        //  LƯU USER PERMISSIONS (tùy chỉnh riêng)
        // =====================================================
        public void SaveUserPermissions(int userId, List<UserPermission> perms)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();

            foreach (var p in perms)
            {
                var cmd = new SqlCommand(@"
                    IF EXISTS (SELECT 1 FROM User_Permissions WHERE User_ID=@uid AND Module_ID=@mid)
                        UPDATE User_Permissions SET
                            Can_View=@v, Can_Create=@c, Can_Edit=@e, Can_Delete=@d, Can_Export=@x
                        WHERE User_ID=@uid AND Module_ID=@mid
                    ELSE
                        INSERT INTO User_Permissions (User_ID, Module_ID, Can_View, Can_Create, Can_Edit, Can_Delete, Can_Export)
                        VALUES (@uid, @mid, @v, @c, @e, @d, @x)", conn);

                cmd.Parameters.AddWithValue("@uid", userId);
                cmd.Parameters.AddWithValue("@mid", p.Module_ID);
                cmd.Parameters.AddWithValue("@v", p.Can_View);
                cmd.Parameters.AddWithValue("@c", p.Can_Create);
                cmd.Parameters.AddWithValue("@e", p.Can_Edit);
                cmd.Parameters.AddWithValue("@d", p.Can_Delete);
                cmd.Parameters.AddWithValue("@x", p.Can_Export);
                cmd.ExecuteNonQuery();
            }
        }

        // =====================================================
        //  XÓA USER PERMISSIONS tùy chỉnh (reset về Role mặc định)
        // =====================================================
        public void ResetUserPermissions(int userId)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            new SqlCommand($"DELETE FROM User_Permissions WHERE User_ID = {userId}", conn).ExecuteNonQuery();
        }

        // =====================================================
        //  DELETE USER (soft delete)
        // =====================================================
        public void DeactivateUser(int userId, string modifiedBy)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                UPDATE Users SET Is_Active=0, Modified_Date=GETDATE(), Modified_By=@modBy
                WHERE User_ID=@userId", conn);
            cmd.Parameters.AddWithValue("@modBy", modifiedBy);
            cmd.Parameters.AddWithValue("@userId", userId);
            cmd.ExecuteNonQuery();
        }


        // =====================================================
        //  LẤY QUYỀN CHI TIẾT THEO MODULE:ACTION
        // =====================================================
        public Dictionary<string, bool> GetDetailedPermissions(int userId)
        {
            var result = new Dictionary<string, bool>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                SELECT Module_Code, Action_Name, Is_Granted
                FROM User_Detailed_Permissions
                WHERE User_ID = @uid", conn);
            cmd.Parameters.AddWithValue("@uid", userId);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                string key = r["Module_Code"].ToString() + ":" + r["Action_Name"].ToString();
                result[key] = Convert.ToBoolean(r["Is_Granted"]);
            }
            return result;
        }

        // =====================================================
        //  LƯU QUYỀN CHI TIẾT THEO MODULE:ACTION
        // =====================================================
        public void SaveDetailedPermissions(int userId, Dictionary<string, bool> perms, string changedBy)
        {
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            using var tran = conn.BeginTransaction();
            try
            {
                // Xóa quyền cũ
                var delCmd = new SqlCommand(
                    "DELETE FROM User_Detailed_Permissions WHERE User_ID = @uid", conn, tran);
                delCmd.Parameters.AddWithValue("@uid", userId);
                delCmd.ExecuteNonQuery();

                // Insert quyền mới
                foreach (var kv in perms)
                {
                    var parts = kv.Key.Split(':', 2);
                    if (parts.Length != 2) continue;

                    var insCmd = new SqlCommand(@"
                        INSERT INTO User_Detailed_Permissions
                            (User_ID, Module_Code, Action_Name, Is_Granted, Updated_By, Updated_At)
                        VALUES
                            (@uid, @mod, @act, @granted, @by, GETDATE())", conn, tran);
                    insCmd.Parameters.AddWithValue("@uid", userId);
                    insCmd.Parameters.AddWithValue("@mod", parts[0]);
                    insCmd.Parameters.AddWithValue("@act", parts[1]);
                    insCmd.Parameters.AddWithValue("@granted", kv.Value);
                    insCmd.Parameters.AddWithValue("@by", changedBy);
                    insCmd.ExecuteNonQuery();
                }

                // Ghi audit log
                string detail = $"Cập nhật {perms.Count} quyền chi tiết";
                var logCmd = new SqlCommand(@"
                    INSERT INTO Permission_Audit_Log (User_ID, Changed_By, Action_Detail, Changed_At)
                    VALUES (@uid, @by, @detail, GETDATE())", conn, tran);
                logCmd.Parameters.AddWithValue("@uid", userId);
                logCmd.Parameters.AddWithValue("@by", changedBy);
                logCmd.Parameters.AddWithValue("@detail", detail);
                logCmd.ExecuteNonQuery();

                tran.Commit();
            }
            catch
            {
                tran.Rollback();
                throw;
            }
        }

        // =====================================================
        //  LẤY TEMPLATE QUYỀN MẶC ĐỊNH THEO ROLE
        // =====================================================
        public Dictionary<string, bool> GetRolePermissionTemplate(int roleId)
        {
            var result = new Dictionary<string, bool>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                SELECT Module_Code, Action_Name, Is_Granted
                FROM Role_Detailed_Permissions
                WHERE Role_ID = @rid", conn);
            cmd.Parameters.AddWithValue("@rid", roleId);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                string key = r["Module_Code"].ToString() + ":" + r["Action_Name"].ToString();
                result[key] = Convert.ToBoolean(r["Is_Granted"]);
            }
            return result;
        }

        // =====================================================
        //  LẤY LỊCH SỬ THAY ĐỔI QUYỀN (AUDIT LOG)
        // =====================================================
        public List<PermissionAuditLog> GetPermissionAuditLog(int userId)
        {
            var result = new List<PermissionAuditLog>();
            using var conn = DatabaseHelper.GetConnection();
            conn.Open();
            var cmd = new SqlCommand(@"
                SELECT TOP 50 Changed_At, Changed_By, Action_Detail
                FROM Permission_Audit_Log
                WHERE User_ID = @uid
                ORDER BY Changed_At DESC", conn);
            cmd.Parameters.AddWithValue("@uid", userId);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                result.Add(new PermissionAuditLog
                {
                    Changed_At = Convert.ToDateTime(r["Changed_At"]),
                    Changed_By = r["Changed_By"]?.ToString() ?? "",
                    Action_Detail = r["Action_Detail"]?.ToString() ?? ""
                });
            }
            return result;
        }
        // =====================================================
        //  MAP
        // =====================================================
        private AppUser MapUser(SqlDataReader r) => new AppUser
        {
            User_ID = Convert.ToInt32(r["User_ID"]),
            Username = r["Username"]?.ToString() ?? "",
            Full_Name = r["Full_Name"]?.ToString() ?? "",
            Email = r["Email"]?.ToString() ?? "",
            Phone = r["Phone"]?.ToString() ?? "",
            Department = r["Department"]?.ToString() ?? "",
            Role_ID = Convert.ToInt32(r["Role_ID"]),
            Role_Name = r["Role_Name"]?.ToString() ?? "",
            Is_Active = Convert.ToBoolean(r["Is_Active"]),
            Must_Change_Password = Convert.ToBoolean(r["Must_Change_Password"]),
            Last_Login = r["Last_Login"] != DBNull.Value ? Convert.ToDateTime(r["Last_Login"]) : null,
            Created_Date = r["Created_Date"] != DBNull.Value ? Convert.ToDateTime(r["Created_Date"]) : null,
            Created_By = r["Created_By"]?.ToString() ?? ""
        };
    }
}