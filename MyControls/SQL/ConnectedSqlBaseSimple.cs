using System;
using System.Transactions;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using System.Collections.Generic;

namespace MyControls
{
    public abstract class ConnectedSqlBaseSimple : IDisposable
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
        public virtual async Task<int> NewExecuteQueryAsync(string QueryString)
        {
            // 2019/12/09 성공은 했고 어차피 한 문장이니 일단 트랜잭션 적용 보류
            //using (SqlTransaction tr = conn.BeginTransaction())
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    //cmd.Transaction = tr;
                    int returnInt = await cmd.ExecuteNonQueryAsync();
                    //if (!returnInt.Equals(-1)) await Task.Run(() => tr.Commit());
                    //else await Task.Run(() => tr.Rollback());
                    return returnInt;
                }
                catch (Exception ex)
                {
                    //try
                    //{
                    //    //await Task.Run(() => tr.Rollback());
                    //}
                    //catch (Exception ex2)
                    //{
                    //    throw ex2;
                    //}
                    System.Diagnostics.Debug.Print(ex.ToString());
                    return -1;
                }
            }
        }
        public virtual async Task<bool> ExecuteQueryAsync(string QueryString)
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
        public async Task<string> BulkUploadToGlobalTempTableAsync(string[] array, string columnName, bool isSql2000 = false)
        {
            using (DataTable tempTable = new DataTable())
            {
                tempTable.Columns.Add(columnName);
                foreach (string reportNumber in array)
                {
                    DataRow dr = tempTable.NewRow();
                    dr[columnName] = reportNumber;
                    tempTable.Rows.Add(dr);
                }
                return await BulkUploadToGlobalTempTableAsync(tempTable, isSql2000);
            }
        }
        public async Task<string> BulkUploadToGlobalTempTableAsync(List<string> list, string columnName, bool isSql2000 = false)
        {
            using (DataTable tempTable = new DataTable())
            {
                tempTable.Columns.Add(columnName);
                foreach (string reportNumber in list)
                {
                    DataRow dr = tempTable.NewRow();
                    dr[columnName] = reportNumber;
                    tempTable.Rows.Add(dr);
                }
                return await BulkUploadToGlobalTempTableAsync(tempTable, isSql2000);
            }
        }
        public async Task<string> BulkUploadToGlobalTempTableAsync(DataTable dt, bool isSql2000 = false)
        {
            StringBuilder columnDefinition = new StringBuilder();
            foreach (DataColumn dc in dt.Columns)
            {
                string typeString = string.Empty;
                if (dc.DataType.Equals(typeof(string))) typeString = isSql2000? "varchar(2000)" : "varchar(MAX)";
                else if (dc.DataType.Equals(typeof(bool))) typeString = "bit";
                else if (dc.DataType.Equals(typeof(byte))) typeString = "tinyint";
                else if (dc.DataType.Equals(typeof(short))) typeString = "smallint";
                else if (dc.DataType.Equals(typeof(int))) typeString = "int";
                else if (dc.DataType.Equals(typeof(long))) typeString = "bigint";
                else if (dc.DataType.Equals(typeof(decimal))) typeString = "numeric(18, 4)";
                else if (dc.DataType.Equals(typeof(float))) typeString = "float";
                else if (dc.DataType.Equals(typeof(double))) typeString = "float";
                else if (dc.DataType.Equals(typeof(DateTime))) typeString = "datetime";
                else throw new NotImplementedException();
                columnDefinition.AppendFormat("[{0}] {1}, ", dc.ColumnName, typeString);
            }
            columnDefinition.Remove(columnDefinition.Length - 2, 2);
            try
            {
                string tempTable = await CreateGlobalTempTableAsync(columnDefinition.ToString());
                if (!(tempTable is default(string)))
                {
                    using (SqlBulkCopy bulkcopy = new SqlBulkCopy(conn)
                    {
                        DestinationTableName = tempTable
                    })
                    {
                        await bulkcopy.WriteToServerAsync(dt);
                    }
                    return tempTable;
                }
                else return default;
            }
            catch
            {
                return default;
            }
        }
        public async Task<string> BulkUploadToGlobalTempTableAsync(DataTable dt, string columnDefinition)
        {
            try
            {
                string tempTable = await CreateGlobalTempTableAsync(columnDefinition);
                if (!(tempTable is default(string)))
                {
                    using (SqlBulkCopy bulkcopy = new SqlBulkCopy(conn)
                    {
                        DestinationTableName = tempTable
                    })
                    {
                        await bulkcopy.WriteToServerAsync(dt);
                    }
                    return tempTable;
                }
                else return default(string);
            }
            catch
            {
                return default(string);
            }
        }
        public async Task<string> CreateGlobalTempTableAsync(string columnDefinition)
        {
            string tempTable = string.Format("##{0}", System.IO.Path.GetRandomFileName().Replace(".", ""));
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = string.Format("Create Table {0} ({1});", tempTable, columnDefinition);
                    await cmd.ExecuteNonQueryAsync();
                    return tempTable;
                }
                catch
                {
                    return default(string);
                }
            }
        }
        public async Task<bool> DropTableAsync(string tableName)
        {
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = String.Format("Drop Table {0};", tableName);
                    switch(await cmd.ExecuteNonQueryAsync())
                    {
                        case -1:
                            return false;
                        case 0:
                            return false;
                        default:
                            return true;
                    }
                }
                catch
                {
                    return false;
                }
            }
        }
        public async Task<bool> BulkUploadAsync(string destTable, DataTable dt)
        {
            try
            {
                using (SqlBulkCopy bulkcopy = new SqlBulkCopy(conn)
                {
                    DestinationTableName = destTable
                })
                {
                    await bulkcopy.WriteToServerAsync(dt);
                }
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
                using (SqlBulkCopy bulkcopy = new SqlBulkCopy(conn)
                {
                    DestinationTableName = destTable
                })
                {
                    foreach (SqlBulkCopyColumnMapping map in sqlBulkCopyColumnMappingCollection)
                    {
                        bulkcopy.ColumnMappings.Add(map);
                    }
                    await bulkcopy.WriteToServerAsync(dt);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }
        public async Task<DataTable> ReturnDTProcedureAsync(string ProcedureName, List<KeyValuePair<string, string>> list)
        {
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
                        using(DataTable modTable = new DataTable())
                        {
                            modTable.Load(dr);
                            return modTable;
                        }
                    }
                }
                catch
                {
                    return null;
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
        public async Task<object> ReturnScalarObjectAsync(string QueryString, int timeout = 30)
        {
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    cmd.CommandTimeout = timeout;
                    return await cmd.ExecuteScalarAsync();
                }
                catch
                {
                    return default(object);
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
                    using (SqlDataReader rd = await cmd.ExecuteReaderAsync())
                    {
                        while (rd.Read())
                        {
                            list.Add(rd.GetString(0));
                        }
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
                    using (SqlDataReader rd = await cmd.ExecuteReaderAsync())
                    {
                        while (rd.Read())
                        {
                            list.Add((T)rd.GetValue(0));
                        }
                    }
                    return list;
                }
                catch
                {
                    return default(List<T>);
                }
            }
        }
        public virtual async Task<DataTable> ReturnDTAsync(string QueryString, int Timeout = 30)
        {
            using (SqlCommand cmd = new SqlCommand())
            {
                try
                {
                    cmd.Connection = conn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = QueryString;
                    cmd.CommandTimeout = Timeout;
                    using (SqlDataReader dr = await cmd.ExecuteReaderAsync())
                    {
            
                        using(DataTable dt = new DataTable())
                        {
                            dt.Load(dr);
                            return dt;
                        }
                    }
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
        ~ConnectedSqlBaseSimple()
        {
            conn.Close();
            Dispose(false);
        }
    }
}
