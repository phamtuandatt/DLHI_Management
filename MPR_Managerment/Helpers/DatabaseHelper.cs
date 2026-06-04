using Microsoft.Data.SqlClient;

namespace MPR_Managerment.Helpers
{
    public static class DatabaseHelper
    {
        private static readonly string _connectionString =
        "Server=tcp:dlhivietnam.database.windows.net,1433;Initial Catalog=MPR_Management;User ID=DLHI_Admin;Password=Hoangquyen@1905;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;";
        //"Server=192.168.88.128,1433;Database=MPR_Managerment_03062026;User Id=sa;Password=Hoangquyen@1905;TrustServerCertificate=True;Encrypt=Optional;";
        //"Server=DESKTOP-KD2BPDJ;Initial Catalog=MPR_Managerment_03062026;User ID=sa;Password=Aa123456@;Encrypt=True;TrustServerCertificate=true;Connection Timeout=30;";
        //"Server=DATPC;Initial Catalog=MPR_Managerment_03062026;User ID=sa;Password=Aa123456@;Encrypt=True;TrustServerCertificate=true;Connection Timeout=30;";


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