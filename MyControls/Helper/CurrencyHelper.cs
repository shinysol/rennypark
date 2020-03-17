using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls.Helper
{
    public static class CurrencyHelper
    {
        public static string ToCurrencyKRW(this decimal amnt)
        {
            return ToCurrencyKRW<decimal>(amnt);
        }
        private static string ToCurrencyKRW<T>(T amnt)
        {
            bool negative = false;
            if (amnt.ToString().Substring(0, 1).Equals("-"))
            {
                negative = true;
            }
            string amntStr = amnt.ToString().Replace("-", "");
            int lth = amntStr.Length;
            int dotPos = amntStr.IndexOf(".");
            if (!dotPos.Equals(-1))
            {
                lth = dotPos;
            }
            if (lth <= 3)
            {
                return amntStr;
            }
            int pos = lth % 3;
            StringBuilder str = new StringBuilder();
            if (pos != 0)
            {
                str.Append(amntStr.Remove(pos));
                str.Append(",");
            }
            while (pos < lth)
            {
                str.Append(amntStr.Substring(pos, 3));
                str.Append(",");
                pos += 3;
            }
            str.Remove(str.Length - 1, 1);
            if (!dotPos.Equals(-1))
            {
                str.Append(amntStr.Substring(dotPos));
            }
            //str.Insert(0, "\\ ");
            if (negative)
            {
                str = str.Insert(0, "-");
            }
            return str.ToString();
        }
    }
}
