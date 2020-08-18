using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MyControls.Helper
{
    public static class RegexHelper
    {
        public static bool IsDigit(this string str, int digits)
        {
            return Regex.IsMatch(str, $@"\d{{{digits}}}");
        }
        public static bool IsAlphaNumeric(this string str)
        {
            return Regex.IsMatch(str, @"^[A-Z0-9]+$");
        }
        public static bool IsCargoManagementNumberType(this string str)
        {
            return Regex.IsMatch(str, @"^\d{2}[a-zA-Z0-9]{9}(-\d{4})?(-\d{4})?$");
        }
    }
}
