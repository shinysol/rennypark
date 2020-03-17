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
        public static bool IsNullOrEmpty(this DataTable dt)
        {
            return dt is null || dt.Rows.Count.Equals(0);
        }
        public static bool IsNullOrEmpty(this DataView dv)
        {
            return dv is null || dv.Table.Rows.Count.Equals(0);
        }
        
        
    }
}
