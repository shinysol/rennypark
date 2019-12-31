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
    public class GeneralSqlConnected : ConnectedSqlBaseSimple
    {
        public GeneralSqlConnected(string DataSource, string InitialCatalog, string UserId, string UserPassword, int ConnectTimeOut = 5)
        {
            SqlCsBuilder.DataSource = DataSource;
            SqlCsBuilder.InitialCatalog = InitialCatalog;
            SqlCsBuilder.UserID = UserId;
            SqlCsBuilder.Password = UserPassword;
            SqlCsBuilder.ConnectTimeout = ConnectTimeOut;
            conn = new SqlConnection(SqlCsBuilder.ToString());
            try
            {
                conn.Open();
            }
            catch(Exception ex)
            {
                if(ex is TimeoutException)
                {
                    throw new TimeoutException("Connection Timeout");
                }
                else
                {
                    throw new Exception("Error Occured.");
                }
            }
        }
    }
}