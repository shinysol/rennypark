using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace MyControls.Helper
{
    public static class DataHelper
    {
        public static bool IsNull(this DataTable dt)
        {
            return dt is null;
        }
        public static bool IsNull(this DataView dv)
        {
            return dv is null;
        }
        public static bool IsNullOrEmpty(this DataTable dt)
        {
            return dt is null || dt.Rows.Count.Equals(0);
        }
        public static bool IsNullOrEmpty(this DataView dv)
        {
            return dv is null || dv.Table.Rows.Count.Equals(0);
        }
        //public static void ColumnReplace(ref DataTable dt, string columnName, Dictionary<string, string> dic)
        //{
        //    if (!dt.Columns.Contains(columnName)) throw new InvalidOperationException($"There is no column named {columnName}.");
        //    if (!dt.Columns[columnName].DataType.Equals(typeof(string))) throw new InvalidOperationException("This operation is only with string type columns.");
        //    foreach (DataRow dr in dt.Rows)
        //    {
        //        dr.SetField(columnName, dic.Aggregate(dr[columnName].ToString(), (result, s) => result.Replace(s.Key, s.Value)));
        //    }
        //}
        
    }
}
