using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls.Helper
{
    public static class EnumerationHelper
    {
        public static string ToString(this List<string> list)
        {
            return string.Format(", ", list);
        }
    }
}
