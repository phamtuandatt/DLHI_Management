// ============================================================
//  FILE: Models/UserModels.cs
//  Chứa toàn bộ Models liên quan đến User & Permission
// ============================================================

using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace MPR_Managerment.Models
{
    // ===== USER =====
    public class AppUser
    {
        public int User_ID { get; set; }
        public string Username { get; set; } = "";
        public string Full_Name { get; set; } = "";
        public string Email { get; set; } = "";
        public string Phone { get; set; } = "";
        public string Department { get; set; } = "";
        public int Role_ID { get; set; }
        public string Role_Name { get; set; } = "";
        public bool Is_Active { get; set; } = true;
        public bool Must_Change_Password { get; set; } = false;
        public DateTime? Last_Login { get; set; }
        public DateTime? Created_Date { get; set; }
        public string Created_By { get; set; } = "";
    }

    // ===== ROLE =====
    public class Role
    {
        public int Role_ID { get; set; }
        public string Role_Name { get; set; } = "";
        public string Description { get; set; } = "";
        public bool Is_Active { get; set; } = true;
    }

    // ===== MODULE =====
    public class AppModule
    {
        public int Module_ID { get; set; }
        public string Module_Code { get; set; } = "";
        public string Module_Name { get; set; } = "";
        public int Sort_Order { get; set; }
    }

    // ===== PERMISSION (quyền hiệu lực cuối) =====
    public class UserPermission
    {
        public int Module_ID { get; set; }
        public string Module_Code { get; set; } = "";
        public string Module_Name { get; set; } = "";
        public bool Can_View { get; set; }
        public bool Can_Create { get; set; }
        public bool Can_Edit { get; set; }
        public bool Can_Delete { get; set; }
        public bool Can_Export { get; set; }
        public bool Is_Custom_Override { get; set; }
    }

    // ===== LOGIN RESULT =====
    public class LoginResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = "";
        public AppUser? User { get; set; }
        public List<UserPermission> Permissions { get; set; } = new();
    }

    // ===== SESSION — lưu thông tin user đang đăng nhập =====
    public static class AppSession
    {
        public static AppUser? CurrentUser { get; set; }
        public static List<UserPermission> Permissions { get; set; } = new();

        public static bool IsLoggedIn => CurrentUser != null;
        public static bool IsAdmin => CurrentUser?.Role_Name == "Admin";

        /// <summary>Lấy quyền của module theo code (PROJECT, PO, MPR...)</summary>
        public static UserPermission? GetPermission(string moduleCode)
            => Permissions.Find(p => p.Module_Code == moduleCode);

        public static bool CanView(string moduleCode)
            => IsAdmin || (GetPermission(moduleCode)?.Can_View ?? false);

        public static bool CanCreate(string moduleCode)
            => IsAdmin || (GetPermission(moduleCode)?.Can_Create ?? false);

        public static bool CanEdit(string moduleCode)
            => IsAdmin || (GetPermission(moduleCode)?.Can_Edit ?? false);

        public static bool CanDelete(string moduleCode)
            => IsAdmin || (GetPermission(moduleCode)?.Can_Delete ?? false);

        public static bool CanExport(string moduleCode)
            => IsAdmin || (GetPermission(moduleCode)?.Can_Export ?? false);

        public static void Clear()
        {
            CurrentUser = null;
            Permissions.Clear();
        }
    }
}

