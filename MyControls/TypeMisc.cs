using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls
{
    static class TypeMisc
    {
        public static bool IsTextAllowed(string text)
        {
            if (text.Equals(string.Empty))
            {
                return false;
            }
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("[^0-9.-]+"); //regex that matches disallowed text
            return !regex.IsMatch(text);
        }
    }
}
