using System;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls.Helper
{
    public static class ObjectHelper
    {
        public static bool ToBool(this object obj)
        {
            if (obj is CheckBox)
            {
                return (bool)(obj as CheckBox).IsChecked;
            }
            if (obj is null || obj.Equals(DBNull.Value))
            {
                return false;
            }
            return (bool)obj;
        }
        public static bool EqualsDBNull(this object obj, bool TrueIfNull = true)
        {
            if (obj is null) return TrueIfNull;
            return obj.Equals(DBNull.Value);
        }
        public static string ToCurrency(this object obj, bool withoutDecimalPart = false)
        {
            if (obj is null || obj.EqualsDBNull()) return string.Empty;
            return withoutDecimalPart ? double.Parse(obj.ToString()).ToString("0") : double.Parse(obj.ToString()).ToString("0.##");
        }
    }
}
