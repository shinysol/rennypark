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
        public static bool Equals(this string[] target, string[] source, int sourcePosition)
        {
            for (int i = 0; i < target.Count(); i++)
            {
                if (!target[i].Equals(source[sourcePosition + i])) return false;
            }
            return true;
        }

        public static string Split(this string str, int limit, int order)
        {
            if (order < 0) return string.Empty;
            return str.Length <= limit * order ? string.Empty : str.Substring(limit * order, Math.Min(str.Length - limit * order, 75));
        }
        public static string ToString(this string[] strArray, string separator)
        {
            return string.Join(separator, strArray);
        }

        public static string GetUntil(this string str, string value)
        {
            if (str.IndexOf(value, StringComparison.Ordinal) > 0)
            {
                return str.Substring(0, str.IndexOf(value, StringComparison.Ordinal));
            }
            return string.Empty;
        }
        public static string GetAfter(this string str, string value)
        {
            if (str.IndexOf(value, StringComparison.Ordinal) > 0)
            {
                return str.Substring(str.IndexOf(value, StringComparison.Ordinal) + value.Length);
            }
            return string.Empty;
        }
        public static string GetUntilFromEnd(this string str, string value)
        {
            if (str.LastIndexOf(value, StringComparison.Ordinal) > 0)
            {
                return str.Substring(str.LastIndexOf(value, StringComparison.Ordinal) + value.Length);
            }
            return string.Empty;
        }
        public static string GetAfterFromEnd(this string str, string value)
        {
            if (str.LastIndexOf(value, StringComparison.Ordinal) > 0)
            {
                return str.Substring(0, str.LastIndexOf(value, StringComparison.Ordinal) + value.Length - 1);
            }
            return string.Empty;
        }
    }
}
