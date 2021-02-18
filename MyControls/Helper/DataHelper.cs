using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Reflection;

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

        public static DataTable ToDataTable<T>(this List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties  
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Defining type of data column gives proper data table   
                var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                //Setting column names as Property names  
                dataTable.Columns.Add(prop.Name, type);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows  
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable  
            return dataTable;
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
