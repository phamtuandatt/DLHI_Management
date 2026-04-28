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
        public async Task<DataTable> GetProducts()
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand("SELECT *FROM Products", conn);
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
        public async Task<bool> SaveProduct_Async(ProductModel product, bool isUpdate)
        {
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                // Chọn Procedure tương ứng
                string procName = isUpdate ? "sp_UpdateProduct" : "sp_InsertProduct";
                SqlCommand cmd = new SqlCommand(procName, conn);
                cmd.CommandType = CommandType.StoredProcedure;

                // Truyền tham số
                if (isUpdate) cmd.Parameters.AddWithValue("@id", product.Id);

                cmd.Parameters.AddWithValue("@name", product.Name ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@des_2", product.Des2 ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@code", product.Code ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@prod_material_code", product.ProdMaterialCode ?? (object)DBNull.Value);

                // Các thông số kỹ thuật (Decimal)
                cmd.Parameters.AddWithValue("@a_thinkness", product.A_Thickness ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@b_depth", product.B_Depth ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@c_witdth", product.C_Width ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@d_web", product.D_Web ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@e_flag", product.E_Flag ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@f_length", product.F_Length ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@g_weight", product.G_Weight ?? (object)DBNull.Value);

                cmd.Parameters.AddWithValue("@used_note", product.UsedNote ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@prod_origin_id", product.ProdOriginId);
                cmd.Parameters.AddWithValue("@prod_standard_id", product.ProdStandardId);
                cmd.Parameters.AddWithValue("@prod_material_cate_id", product.ProdMaterialCateId);
                cmd.Parameters.AddWithValue("@prod_material_id", product.ProdMaterialId);
                cmd.Parameters.AddWithValue("@prod_material_detail_id", product.ProdMaterialDetailId);

                try
                {
                    await conn.OpenAsync();
                    int rows = await cmd.ExecuteNonQueryAsync();
                    return rows > 0;
                }
                catch (Exception ex)
                {
                    throw new Exception("Lỗi Database: " + ex.Message);
                }
            }
        }

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

        public async Task<string> GetCodeExistedByMaterilDetail(int materialId)
        {
            string sqlQuery = string.Format("SELECT TOP 1 item_code_existed FROM Material_Detail WHERE material_detail_code = {0} ORDER BY material_detail_number DESC", materialId);
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand(sqlQuery, conn);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                var itemCOde = "";
                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                    foreach (DataRow row in dt.Rows)
                    {
                        itemCOde = row[0].ToString().Trim();
                    }
                }
                return itemCOde;
            }
        }

        public async Task<DataTable> GetitemExistedList(int materialId)
        {
            string sqlQuery = string.Format("SELECT * FROM Material_Detail WHERE material_detail_code = {0} ORDER BY material_detail_number DESC", materialId);
            using (SqlConnection conn = DatabaseHelper.GetConnection())
            {
                SqlCommand cmd = new SqlCommand(sqlQuery, conn);

                DataTable dt = new DataTable();
                await conn.OpenAsync(); // Mở kết nối ngầm

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync()) // Đọc dữ liệu ngầm
                {
                    dt.Load(reader);
                }
                return dt;
            }
        }

        public async Task<int> InsertMaterialTypeDetailItem(Material_Detail item)
        {
            string sqlQuery = string.Format("INSERT INTO Material_Detail (material_detail_number, material_detail_name, material_detail_code, item_code_existed) " +
                "VALUES ('{0}', '{1}', '{2}', '{3}')", item.Material_Detail_Number, item.Material_Detail_Name, item.MaterialID, item.Item_Code_Existed);

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
