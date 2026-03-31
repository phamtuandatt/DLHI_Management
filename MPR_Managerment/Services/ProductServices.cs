using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using MPR_Managerment.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace MPR_Managerment.Services
{
    public class ProductServices
    {
        public async Task<DataTable> GetMaterialCates()
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand("SELECT *FROM Material_Categories", conn);
                //cmd.Parameters.AddWithValue("@mprId", mprID);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<DataTable> GetMaterials(int cateId)
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand("SELECT material_id, material_code, material_name FROM Materials WHERE cat_id = @catId", conn);
                cmd.Parameters.AddWithValue("@catId", cateId);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<DataTable> GetOriginals()
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand("SELECT *FROM Origins", conn);
                //cmd.Parameters.AddWithValue("@catId", cateId);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<DataTable> GetStandards()
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand("SELECT *FROM Standards", conn);
                //cmd.Parameters.AddWithValue("@catId", cateId);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<string> GetItemNumberOfMaterialType(int materialId)
        {
            string sqlQuery = string.Format("EXEC GET_ITEM_NUMBER_OF_MATERIAL_TYPE '{0}'", materialId);
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand(sqlQuery, conn);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                var itemNumber = "";
                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                    foreach (DataRow row in dt.Rows)
                    {
                        itemNumber = row[0].ToString().Trim();
                    }
                }
                return itemNumber;
            }
        }

        public async Task<int> InsertMaterialTypeDetailItem(Material_Detail item)
        {
            string sqlQuery = string.Format("INSERT INTO Material_Detail (material_detail_id, material_detail_number, material_detail_name, material_detail_code) " +
                "VALUES ({0}, '{1}', '{2}', '{3}')", item.Material_Detail_Id, item.Material_Detail_Number, item.Material_Detail_Name, item.Material_Detail_Code);

            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(sqlQuery, conn);
                var r = cmd.ExecuteReader();
                if (r.Read()) 
                    return Convert.ToInt32(r["material_detail_id"]);
                return 0;
            }
        }

        public async Task<int> InsertProduct(ProductAddModel item)
        {
            string sqlQuery = string.Format("INSERT INTO Products(code) VALUES ('{0}')", item.Code);

            using (var conn = DatabaseHelper.GetConnection())
            {
                conn.Open();
                var cmd = new SqlCommand(sqlQuery, conn);
                var r = cmd.ExecuteReader();
                if (r.Read())
                    return Convert.ToInt32(r["ID"]);
                return 0;
            }
        }
    }
}
