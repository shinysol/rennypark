using System;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using System.Collections.Generic;

namespace MyControls
{
    public abstract class ConnectedSqlBase : IDisposable
    {
        bool disposed = false;
        SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);  // Dispose
        protected SqlConnection conn = new SqlConnection();
        protected SqlConnectionStringBuilder SqlCsBuilder = new SqlConnectionStringBuilder();   // 통관DB
        public bool ExecuteQuery(string QueryString)
        {
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
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
            }
        }
        public async Task<bool> ExecuteQueryAsync(string QueryString)
        {
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
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
            }
        }
        public bool ExecuteProcedure(string ProcedureName)
        {

            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                try
                {
                    cmd.ExecuteNonQuery();
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
        public bool ExecuteProcedure(string ProcedureName, List<KeyValuePair<string, string>> list)
        {

            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {

                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                foreach (KeyValuePair<string, string> li in list)
                {
                    cmd.Parameters.Add(li.Key, SqlDbType.Char).Value = li.Value;
                }
                try
                {

                    cmd.ExecuteNonQuery();
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
        public bool ExecuteProcedure(string ProcedureName, List<Structs.SqlVar> var)
        {
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
                    cmd.ExecuteNonQuery();
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
        public async Task<bool> ExecuteProcedureAsync(string ProcedureName)
        {

            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                try
                {
                    await cmd.ExecuteNonQueryAsync();
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
        public async Task<bool> ExecuteProcedureAsync(string ProcedureName, List<KeyValuePair<string, string>> list)
        {
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                foreach (KeyValuePair<string, string> li in list)
                {
                    cmd.Parameters.Add(li.Key, SqlDbType.Char).Value = li.Value;
                }
                try
                {

                    await cmd.ExecuteNonQueryAsync();
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }
        public async Task<bool> ExecuteProcedureAsync(string ProcedureName, List<Structs.SqlVar> var)
        {
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                foreach (Structs.SqlVar v in var)
                {
                    object varValue;
                    if (v.varType.Equals(SqlDbType.Udt))
                    {
                        // 180710 UDT 추가
                        SqlParameter sqlParameter = new SqlParameter(v.varNm, v.varType)
                        {
                            UdtTypeName = v.udtNm,
                            Value = (v.varVal as DataTable),
                        };
                        cmd.Parameters.Add(sqlParameter);
                        continue;
                    }
                    if (v.varControl is null)
                    {
                        varValue = v.varVal;
                    }
                    else
                    {
                        varValue = Control.TcmsControl.ControlTextExtractor(v.varControl);
                    }
                    if (v.varSize.Equals(0))
                    {
                        cmd.Parameters.Add(v.varNm, v.varType).Value = varValue;
                    }
                    else
                    {
                        cmd.Parameters.Add(v.varNm, v.varType, v.varSize).Value = varValue;
                    }
                }
                try
                {
                    await cmd.ExecuteNonQueryAsync();
                    return true;
                }
                catch(Exception ex)
                {
                    throw ex;
                }
            }
        }
        public bool BulkUpload(string destTable, DataTable dt)
        {
            try
            {
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn)
                {
                    DestinationTableName = destTable,
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
            try
            {
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
        public async Task<bool> BulkUploadAsync(string destTable, DataTable dt, List<SqlBulkCopyColumnMapping> sqlBulkCopyColumnMappingCollection)
        {
            try
            {
                SqlBulkCopy bulkcopy = new SqlBulkCopy(conn)
                {
                    DestinationTableName = destTable
                };
                foreach (SqlBulkCopyColumnMapping map in sqlBulkCopyColumnMappingCollection)
                {
                    bulkcopy.ColumnMappings.Add(map);
                }
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
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                foreach (KeyValuePair<string, string> li in list)
                {
                    cmd.Parameters.Add(li.Key, SqlDbType.Char).Value = li.Value;
                }
                try
                {

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
            }
        }
        public async Task<DataTable> ReturnDTProcedureAsync(string ProcedureName, List<KeyValuePair<string, string>> list)
        {
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
                foreach (KeyValuePair<string, string> li in list)
                {
                    cmd.Parameters.Add(li.Key, SqlDbType.Char).Value = li.Value;
                }
                try
                {

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
            }
        }
        public T ReturnScalarData<T>(string QueryString) where T : struct
        {
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    return (T)cmd.ExecuteScalar();
                }
                catch
                {
                    return default(T);
                }
            }
        }
        public async Task<string> ReturnScalarStringAsync(string QueryString)
        {
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    return (await cmd.ExecuteScalarAsync()).ToString();
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        public async Task<byte[]> ReturnByteArrayAsync(string queryString)
        {
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    return (byte[])(await cmd.ExecuteScalarAsync());
                }
                catch
                {
                    return null;
                }
            }
        }
        public async Task<T> ReturnScalarDataAsync<T>(string QueryString, int timeout = 30) where T : struct
        {
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    cmd.CommandTimeout = timeout;
                    return (T)(await cmd.ExecuteScalarAsync());
                }
                catch
                {
                    return default(T);
                }
            }
        }
        public async Task<T> ReturnScalarDataAsync<T>(string ProcedureName, List<Structs.SqlVar> var) where T : struct
        {
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
                    return (T)(await cmd.ExecuteScalarAsync());
                }
                catch
                {
                    return default(T);
                }
            }
        }
        public List<string> ReturnStringList(string QueryString)
        {
            List<string> list = new List<string>();
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    SqlDataReader rd = cmd.ExecuteReader();
                    while (rd.Read())
                    {
                        list.Add(rd.GetString(0));
                    }
                    return list;
                }
                catch
                {
                    return default(List<string>);
                }
            }
        }
        public async Task<List<string>> ReturnStringListAsync(string QueryString)
        {
            List<string> list = new List<string>();
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    SqlDataReader rd = await cmd.ExecuteReaderAsync();
                    while (rd.Read())
                    {
                        list.Add(rd.GetString(0));
                    }
                    return list;
                }
                catch
                {
                    return default(List<string>);
                }
            }
        }
        public async Task<List<T>> ReturnScalarListAsync<T>(string QueryString) where T : struct
        {
            List<T> list = new List<T>();
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    SqlDataReader rd = await cmd.ExecuteReaderAsync();
                    while (rd.Read())
                    {
                        list.Add((T)rd.GetValue(0));
                    }
                    return list;
                }
                catch
                {
                    return default(List<T>);
                }
            }
        }
        public DataTable ReturnDT(string QueryString)
        {
            DataTable dt = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    dt.Load(cmd.ExecuteReader());
                    return dt;
                }
                catch
                {
                    return null;
                }
            }
        }
        public DataTable ReturnDT(string QueryString, int Timeout = 30)
        {
            DataTable dt = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    cmd.CommandTimeout = Timeout;
                    SqlDataReader dr = cmd.ExecuteReader();
                    dt.Load(dr);
                    return dt;
                }
                catch
                {
                    return null;
                }
            }
        }
        public async Task<DataTable> ReturnDTAsync(string QueryString, int Timeout = 30)
        {
            DataTable dt = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    cmd.CommandTimeout = Timeout;
                    SqlDataReader dr = await cmd.ExecuteReaderAsync();
                    dt.Load(dr);
                    return dt;
                }
                catch
                {
                    return null;
                }
            }
        }
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
                conn.Dispose();
                //
            }
            // Free any unmanaged objects here.
            //
            disposed = true;
        }
        ~ConnectedSqlBase()
        {
            conn.Close();
            Dispose(false);
        }
    }
}
