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
    public class Sqltotcms : IDisposable
    {
        bool disposed = false;
        SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);  // Dispose
                                                                    // Public implementation of Dispose pattern callable by consumers.
        public SqlConnectionStringBuilder SqlCsBuilder = new SqlConnectionStringBuilder();   // 통관DB
        public SqlConnectionStringBuilder SqlCsBuilder2 = new SqlConnectionStringBuilder();   // 통관DB
        public SqlConnectionStringBuilder sqlCsBuilderReq = new SqlConnectionStringBuilder();  // 에이원DB
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
        ~Sqltotcms()
        {
            Dispose(false);
        }
        public Sqltotcms()
        {
            SqlCsBuilder.DataSource = "14.63.173.204";
            SqlCsBuilder.InitialCatalog = "AONETCMS_HC";
            SqlCsBuilder.UserID = "aonetcms";
            SqlCsBuilder.Password = "aonetcms";
            SqlCsBuilder.ConnectTimeout = 60;
            SqlCsBuilder2.DataSource = "14.63.173.204";
            SqlCsBuilder2.InitialCatalog = "AONETCMS";
            SqlCsBuilder2.UserID = "aonetcms";
            SqlCsBuilder2.Password = "aonetcms";
            SqlCsBuilder2.ConnectTimeout = 60;
            sqlCsBuilderReq.DataSource = "14.63.173.204";
            sqlCsBuilderReq.InitialCatalog = "AONE";
            sqlCsBuilderReq.UserID = "aonetcms";
            sqlCsBuilderReq.Password = "aonetcms";
            sqlCsBuilderReq.ConnectTimeout = 60;
        }
        public bool ExecuteQuery(string queryString, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:

                    conn = new SqlConnection(SqlCsBuilder2.ToString());
                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
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
        public async Task<bool> ExecuteQueryAsync(string queryString, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    await conn.OpenAsync();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
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
        public bool ExecuteProcedure(string ProcedureName, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                switch (dbOrder)
                {
                    case 1:
                        conn = new SqlConnection(sqlCsBuilderReq.ToString());
                        break;
                    case 2:
                        conn = new SqlConnection(SqlCsBuilder2.ToString());

                        break;
                    case 3:
                        conn = new SqlConnection(SqlCsBuilder.ToString());
                        break;
                    default:
                        break;
                }
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
        public bool ExecuteProcedure(string ProcedureName, List<KeyValuePair<string, string>> list, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                switch (dbOrder)
                {
                    case 1:
                        conn = new SqlConnection(sqlCsBuilderReq.ToString());
                        break;
                    case 2:
                        conn = new SqlConnection(SqlCsBuilder2.ToString());

                        break;
                    case 3:
                        conn = new SqlConnection(SqlCsBuilder.ToString());
                        break;
                    default:
                        break;
                }
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
        public bool ExecuteProcedure(string ProcedureName, List<Structs.SqlVar> var, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                switch (dbOrder)
                {
                    case 1:
                        conn = new SqlConnection(sqlCsBuilderReq.ToString());
                        break;
                    case 2:
                        conn = new SqlConnection(SqlCsBuilder2.ToString());

                        break;
                    case 3:
                        conn = new SqlConnection(SqlCsBuilder.ToString());
                        break;
                    default:
                        break;
                }
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
        public async Task<bool> ExecuteProcedureAsync(string ProcedureName, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                switch (dbOrder)
                {
                    case 1:
                        conn = new SqlConnection(sqlCsBuilderReq.ToString());
                        break;
                    case 2:
                        conn = new SqlConnection(SqlCsBuilder2.ToString());

                        break;
                    case 3:
                        conn = new SqlConnection(SqlCsBuilder.ToString());
                        break;
                    default:
                        break;
                }
                cmd.Connection = conn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = ProcedureName;
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
        public async Task<bool> ExecuteProcedureAsync(string ProcedureName, List<KeyValuePair<string, string>> list, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                switch (dbOrder)
                {
                    case 1:
                        conn = new SqlConnection(sqlCsBuilderReq.ToString());
                        break;
                    case 2:
                        conn = new SqlConnection(SqlCsBuilder2.ToString());

                        break;
                    case 3:
                        conn = new SqlConnection(SqlCsBuilder.ToString());
                        break;
                    default:
                        break;
                }
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
                switch (dbOrder)
                {
                    case 1:
                        conn = new SqlConnection(sqlCsBuilderReq.ToString());
                        break;
                    case 2:
                        conn = new SqlConnection(SqlCsBuilder2.ToString());

                        break;
                    case 3:
                        conn = new SqlConnection(SqlCsBuilder.ToString());
                        break;
                    default:
                        break;
                }
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
        public bool BulkUpload(string destTable, DataTable dt, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
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
        public async Task<bool> BulkUploadAsync(string destTable, DataTable dt, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
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
        public DataTable ReturnDTProcedure(string ProcedureName, List<KeyValuePair<string, string>> list, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                switch (dbOrder)
                {
                    case 1:
                        conn = new SqlConnection(sqlCsBuilderReq.ToString());
                        break;
                    case 2:
                        conn = new SqlConnection(SqlCsBuilder2.ToString());

                        break;
                    case 3:
                        conn = new SqlConnection(SqlCsBuilder.ToString());
                        break;
                    default:
                        break;
                }
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
        public async Task<DataTable> ReturnDTProcedureAsync(string ProcedureName, List<KeyValuePair<string, string>> list, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            DataTable modTable = new DataTable();
            using (SqlCommand cmd = new SqlCommand())
            {
                switch (dbOrder)
                {
                    case 1:
                        conn = new SqlConnection(sqlCsBuilderReq.ToString());
                        break;
                    case 2:
                        conn = new SqlConnection(SqlCsBuilder2.ToString());

                        break;
                    case 3:
                        conn = new SqlConnection(SqlCsBuilder.ToString());
                        break;
                    default:
                        break;
                }
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
        public string ReturnString(string queryString, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    return cmd.ExecuteScalar().ToString();
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
        public async Task<string> ReturnStringAsync(string queryString, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    await conn.OpenAsync();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    var st = await cmd.ExecuteScalarAsync();
                    return st.ToString();
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
        public byte ReturnByte(string queryString, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    var result = cmd.ExecuteScalar();
                    if (result != null)
                    {
                        return (byte)cmd.ExecuteScalar();
                    }
                    else
                    {
                        return 0;
                    }
                }
                catch
                {
                    return 255;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public async Task<byte> ReturnByteAsync(string queryString, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    await conn.OpenAsync();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    var result = await cmd.ExecuteScalarAsync();
                    if (result != null)
                    {
                        return (byte)cmd.ExecuteScalar();
                    }
                    else
                    {
                        return 0;
                    }
                }
                catch
                {
                    return 255;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public int ReturnInteger(string queryString, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    var ret = cmd.ExecuteScalar();
                    if (ret != null)
                    {
                        return (int)ret;
                    }
                    return -1;
                }
                catch
                {
                    return -1;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public async Task<int> ReturnIntegerAsync(string queryString, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    await conn.OpenAsync();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    var ret = await cmd.ExecuteScalarAsync();
                    if (ret != null)
                    {
                        return (int)ret;
                    }
                    return -1;
                }
                catch
                {
                    return -1;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
        public double ReturnFloat(string queryString, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    return Convert.ToDouble(cmd.ExecuteScalar());
                }
                catch
                {
                    return 0;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

        }
        public async Task<double> ReturnFloatAsync(string queryString, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    await conn.OpenAsync();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    var dbl = await cmd.ExecuteScalarAsync();
                    return Convert.ToDouble(dbl);
                }
                catch
                {
                    return 0;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }

        }
        public List<int> ReturnIntegerList(string queryString, byte dbOrder)
        {
            List<int> list = new List<int>();
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    SqlDataReader rd = cmd.ExecuteReader();
                    while (rd.Read())
                    {
                        list.Add(rd.GetInt32(0));
                    }
                    return list;
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
        public async Task<List<int>> ReturnIntegerListAsync(string queryString, byte dbOrder)
        {
            List<int> list = new List<int>();
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    await conn.OpenAsync();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    SqlDataReader rd = await cmd.ExecuteReaderAsync();
                    while (rd.Read())
                    {
                        list.Add(rd.GetInt32(0));
                    }
                    return list;
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
        public List<string> ReturnStringList(string queryString, byte dbOrder)
        {
            List<string> list = new List<string>();
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    SqlDataReader rd = cmd.ExecuteReader();
                    while (rd.Read())
                    {
                        list.Add(rd.GetString(0));
                    }
                    return list;
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
        public async Task<List<string>> ReturnStringListAsync(string queryString, byte dbOrder)
        {
            List<string> list = new List<string>();
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());


                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    await conn.OpenAsync();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    SqlDataReader rd = await cmd.ExecuteReaderAsync();
                    while (rd.Read())
                    {
                        list.Add(rd.GetString(0));
                    }
                    return list;
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
        public List<string> ReturnStringListfromByte(string queryString, byte dbOrder)
        {
            List<string> list = new List<string>();
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    SqlDataReader rd = cmd.ExecuteReader();
                    while (rd.Read())
                    {
                        list.Add(rd.GetByte(0).ToString());
                    }
                    return list;
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
        public async Task<List<string>> ReturnStringListfromByteAsync(string queryString, byte dbOrder)
        {
            List<string> list = new List<string>();
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    await conn.OpenAsync();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    SqlDataReader rd = await cmd.ExecuteReaderAsync();
                    while (rd.Read())
                    {
                        list.Add(rd.GetByte(0).ToString());
                    }
                    return list;
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
        public List<KeyValuePair<string, byte>> ReturnPairList(string queryString, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            List<KeyValuePair<String, byte>> list = new List<KeyValuePair<string, byte>>();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())

            {
                try
                {
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    SqlDataReader rd = cmd.ExecuteReader();
                    while (rd.Read())
                    {
                        list.Add(new KeyValuePair<string, byte>(rd.GetString(0), (byte)rd.GetInt16(1)));
                    }
                    return list;
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
        public async Task<List<KeyValuePair<string, byte>>> ReturnPairListAsync(string queryString, byte dbOrder)
        {
            SqlConnection conn = new SqlConnection();
            List<KeyValuePair<String, byte>> list = new List<KeyValuePair<string, byte>>();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())

            {
                try
                {
                    await conn.OpenAsync();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    SqlDataReader rd = await cmd.ExecuteReaderAsync();
                    while (rd.Read())
                    {
                        list.Add(new KeyValuePair<string, byte>(rd.GetString(0), (byte)rd.GetInt16(1)));
                    }
                    return list;
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
        public DataTable ReturnDT(string queryString, byte dbOrder)
        {
            DataTable dt = new DataTable();
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
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
        public DataTable ReturnDT(string tableName, List<string> columns, List<string> where, Dictionary<string, bool> orderBy, byte dbOrder)
        {
            return ReturnDT(SqlMisc.DbSelectString(tableName, columns, where, orderBy), dbOrder);
        }
        public async Task<DataTable> ReturnDTAsync(string queryString, byte dbOrder, int timeout = 30)
        {
            DataTable dt = new DataTable();
            SqlConnection conn = new SqlConnection();
            switch (dbOrder)
            {
                case 1:
                    conn = new SqlConnection(sqlCsBuilderReq.ToString());
                    break;
                case 2:
                    conn = new SqlConnection(SqlCsBuilder2.ToString());

                    break;
                case 3:
                    conn = new SqlConnection(SqlCsBuilder.ToString());
                    break;
                default:
                    break;
            }
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    await conn.OpenAsync();
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = queryString;
                    cmd.CommandTimeout = timeout;
                    SqlDataReader dr = await cmd.ExecuteReaderAsync();
                    dt.Load(dr);
                }
                catch
                {
                    
                }
                finally
                {
                    conn.Close();
                }
                conn.Dispose();
                return dt ?? null;
            }
        }
        public async Task<DataTable> ReturnDTAsync(string tableName, List<string> columns, List<string> where, Dictionary<string,bool> orderBy, byte dbOrder)
        {
            return await ReturnDTAsync(SqlMisc.DbSelectString(tableName,columns,  where, orderBy), dbOrder);
        }
    }
}