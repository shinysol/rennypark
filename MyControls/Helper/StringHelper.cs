using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls.Helper
{
    public static class StringHelper
    {
        public static bool IsNullOrWhiteSpace(this string str)
        {
            return string.IsNullOrWhiteSpace(str);
        }
        public static bool IsNullOrEmpty(this string str)
        {
            return string.IsNullOrEmpty(str);
        }
        public static string JoinWithSeparator(this string[] list, string separator = ", ")
        {
            return string.Join(separator, list.Where(x => !x.IsNullOrWhiteSpace()));
        }
        public static string ForQuery(this string str)
        {
            return str.Replace("'", "").Replace(";", "");
        }
    }
}
