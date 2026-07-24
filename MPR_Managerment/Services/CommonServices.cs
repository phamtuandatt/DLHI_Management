using Microsoft.Data.SqlClient;
using MPR_Managerment.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MPR_Managerment.Services
{
    internal class CommonServices
    {
        /// <summary>
        /// Execute a SQL query and return the result as a DataTable
        /// </summary>
        public DataTable ExecuteQuery(string query, SqlParameter[] parameters = null)
        {
            DataTable dt = new DataTable();
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandTimeout = 60;
                        if (parameters != null)
                            cmd.Parameters.AddRange(parameters);

                        using (var adapter = new SqlDataAdapter(cmd))
                        {
                            adapter.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DB ERROR] {ex.Message}");
                throw;
            }
            return dt;
        }

        /// <summary>
        /// Execute a SQL non-query command (INSERT, UPDATE, DELETE)
        /// </summary>
        public int ExecuteNonQuery(string query, SqlParameter[] parameters = null)
        {
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandTimeout = 60;
                        if (parameters != null)
                            cmd.Parameters.AddRange(parameters);

                        return cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DB ERROR] {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Execute a stored procedure and return the result as a DataTable
        /// </summary>
        public DataTable ExecuteStoredProcedure(string procName, params SqlParameter[] parameters)
        {
            DataTable dt = new DataTable();
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(procName, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 60;
                        if (parameters != null)
                            cmd.Parameters.AddRange(parameters);

                        using (var adapter = new SqlDataAdapter(cmd))
                        {
                            adapter.Fill(dt);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DB ERROR] {ex.Message}");
                throw;
            }
            return dt;
        }

        /// <summary>
        /// Execute a SQL query and return a single scalar value
        /// </summary>
        public object ExecuteScalar(string query, SqlParameter[] parameters = null)
        {
            try
            {
                using (var conn = DatabaseHelper.GetConnection())
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandTimeout = 60;
                        if (parameters != null)
                            cmd.Parameters.AddRange(parameters);

                        return cmd.ExecuteScalar();
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DB ERROR] {ex.Message}");
                throw;
            }
        }
    }
}