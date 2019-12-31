using System;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using System.Collections.Generic;

namespace MyControls
{
    public class GeneralSql : IDisposable
    {
        bool disposed = false;
        SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);  // Dispose
                                                                    // Public implementation of Dispose pattern callable by consumers.
        public SqlConnectionStringBuilder SqlCsBuilder = new SqlConnectionStringBuilder();   // 통관DB
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                handle.Dispose();
                // Free any other managed objects here.
                //
            }
            // Free any unmanaged objects here.
            //
            disposed = true;
        }
        public GeneralSql(string DataSource, string InitialCatalog, string UserId, string UserPassword, int ConnectTimeOut = 60)
        {
            SqlCsBuilder.DataSource = DataSource;
            SqlCsBuilder.InitialCatalog = InitialCatalog;
            SqlCsBuilder.UserID = UserId;
            SqlCsBuilder.Password = UserPassword;
            SqlCsBuilder.ConnectTimeout = ConnectTimeOut;
        }
        public GeneralSql(string UserId, string UserPassword, int ConnectTimeOut = 60)
            : this(string.Empty, string.Empty, UserId, UserPassword, ConnectTimeOut)
        {
            //SqlCsBuilder.DataSource = DataSource;
            //SqlCsBuilder.InitialCatalog = InitialCatalog;
            //SqlCsBuilder.UserID = UserId;
            //SqlCsBuilder.Password = UserPassword;
            //SqlCsBuilder.ConnectTimeout = ConnectTimeOut;
        }
        ~GeneralSql()
        {
            Dispose(false);
        }
        public bool ExecuteQuery(string QueryString)
        {
            SqlConnection conn = new SqlConnection();
            conn = new SqlConnection(SqlCsBuilder.ToString());
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    cmd.ExecuteNonQuery();
                    return true;
                }
                catch
                {
                    return false;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public async Task<bool> ExecuteQueryAsync(string QueryString)
        {
            SqlConnection conn = new SqlConnection();
            conn = new SqlConnection(SqlCsBuilder.ToString());
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    await conn.OpenAsync();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    await cmd.ExecuteNonQueryAsync();
                    return true;
                }
                catch
                {
                    return false;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public bool ExecuteProcedure(string ProcedureName)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                conn = new SqlConnection(SqlCsBuilder.ToString());
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                try
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    return true;
                }
                catch
                {
                    return false;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public bool ExecuteProcedure(string ProcedureName, List<KeyValuePair<string, string>> list)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                conn = new SqlConnection(SqlCsBuilder.ToString());
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                foreach (KeyValuePair<string, string> li in list)
                {
                    cmd.Parameters.Add(li.Key, SqlDbType.Char).Value = li.Value;
                }
                try
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    return true;
                }
                catch
                {
                    return false;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public bool ExecuteProcedure(string ProcedureName, List<Structs.SqlVar> var)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                conn = new SqlConnection(SqlCsBuilder.ToString());
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                foreach (Structs.SqlVar v in var)
                {
                    if (v.varSize.Equals(0))
                    {
                        cmd.Parameters.Add(v.varNm, v.varType).Value = v.varVal;
                    }
                    else
                    {
                        cmd.Parameters.Add(v.varNm, v.varType, v.varSize).Value = v.varVal;
                    }
                }
                try
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    return true;
                }
                catch
                {
                    return false;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public async Task<bool> ExecuteProcedureAsync(string ProcedureName)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                conn = new SqlConnection(SqlCsBuilder.ToString());
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                try
                {
                    conn.Open();
                    await cmd.ExecuteNonQueryAsync();
                    return true;
                }
                catch
                {
                    return false;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public async Task<bool> ExecuteProcedureAsync(string ProcedureName, List<KeyValuePair<string, string>> list)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                conn = new SqlConnection(SqlCsBuilder.ToString());
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                foreach (KeyValuePair<string, string> li in list)
                {
                    cmd.Parameters.Add(li.Key, SqlDbType.Char).Value = li.Value;
                }
                try
                {
                    await conn.OpenAsync();
                    await cmd.ExecuteNonQueryAsync();
                    return true;
                }
                catch
                {
                    return false;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public async Task<bool> ExecuteProcedureAsync(string ProcedureName, List<Structs.SqlVar> var, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                foreach (Structs.SqlVar v in var)
                {
                    if (v.varSize.Equals(0))
                    {
                        cmd.Parameters.Add(v.varNm, v.varType).Value = v.varVal;
                    }
                    else
                    {
                        cmd.Parameters.Add(v.varNm, v.varType, v.varSize).Value = v.varVal;
                    }
                }
                try
                {
                    await conn.OpenAsync();
                    await cmd.ExecuteNonQueryAsync();
                    return true;
                }
                catch
                {
                    return false;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public bool BulkUpload(string destTable, DataTable dt)
        {
            SqlConnection conn = new SqlConnection();
            conn = new SqlConnection(SqlCsBuilder.ToString());
            try
            {
                conn.Open();
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn)
                {
                    DestinationTableName = destTable
                };
                bulkcopy.WriteToServer(dt);
                return true;
            }
            catch
            {
                return false;
            }
        }
        public async Task<bool> BulkUploadAsync(string destTable, DataTable dt)
        {
            SqlConnection conn = new SqlConnection();
            conn = new SqlConnection(SqlCsBuilder.ToString());
            try
            {
                await conn.OpenAsync();
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn)
                {
                    DestinationTableName = destTable
                };
                await bulkcopy.WriteToServerAsync(dt);
                return true;
            }
            catch
            {
                return false;
            }
        }
        public DataTable ReturnDTProcedure(string ProcedureName, List<KeyValuePair<string, string>> list)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                conn = new SqlConnection(SqlCsBuilder.ToString());
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                foreach(KeyValuePair<string, string> li in list)
                {
                    cmd.Parameters.Add(li.Key, SqlDbType.Char).Value = li.Value;
                }
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        modTable.Load(dr);
                    }
                    return modTable;
                }
                catch
                {
                    return null;    
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public async Task<DataTable> ReturnDTProcedureAsync(string ProcedureName, List<KeyValuePair<string, string>> list)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                conn = new SqlConnection(SqlCsBuilder.ToString());
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                foreach (KeyValuePair<string, string> li in list)
                {
                    cmd.Parameters.Add(li.Key, SqlDbType.Char).Value = li.Value;
                }
                try
                {
                    await conn.OpenAsync();
                    using (SqlDataReader dr = await cmd.ExecuteReaderAsync())
                    {
                        modTable.Load(dr);
                    }
                    return modTable;
                }
                catch
                {
                    return null;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public T ReturnScalarData<T>(string QueryString) where T : struct
        {
            SqlConnection conn = new SqlConnection();
            conn = new SqlConnection(SqlCsBuilder.ToString());
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    return (T)cmd.ExecuteScalar();
                }
                catch
                {
                    return default(T);
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public async Task<T> ReturnScalarDataAsync<T>(string QueryString) where T: struct
        {
            SqlConnection conn = new SqlConnection();
            conn = new SqlConnection(SqlCsBuilder.ToString());
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    return (T)(await cmd.ExecuteScalarAsync());
                }
                catch
                {
                    return default(T);
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public DataTable ReturnDT(string queryString)
        {
            DataTable dt = new DataTable();
            SqlConnection conn = new SqlConnection();
            conn = new SqlConnection(SqlCsBuilder.ToString());
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    dt.Load(cmd.ExecuteReader());
                    return dt;
                }
                catch
                {
                    return null;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public async Task<DataTable> ReturnDTAsync(string queryString)
        {
            DataTable dt = new DataTable();
            SqlConnection conn = new SqlConnection();
            conn = new SqlConnection(SqlCsBuilder.ToString());
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    SqlDataReader dr = await cmd.ExecuteReaderAsync();
                    dt.Load(dr);
                    return dt;
                }
                catch
                {
                    return null;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
    }
}