using Microsoft.Data.SqlClient;

namespace MPR_Managerment.Helpers
{
    public static class DatabaseHelper
    {
        private static readonly string _connectionString =
            "Server=tcp:dlhivietnam.database.windows.net,1433;" +
            "Initial Catalog=MPR_Management;" +
            "User ID=DLHI_Admin;" +
            "Password=Hoangquyen@1905;" +
            "Encrypt=True;" +
            "TrustServerCertificate=False;" +
            "Connection Timeout=30;";

        public static SqlConnection GetConnection()
        {
            return new SqlConnection(_connectionString);
        }

        public static bool TestConnection()
        {
            try
            {
                using (var conn = GetConnection())
                {
                    conn.Open();
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}