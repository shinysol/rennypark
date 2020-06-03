using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls.Helper
{
    public static class DateTimeHelper
    {
        public static string To8DigitString(this DateTime dtm) => dtm.ToString("yyyyMMdd");
    }
}
