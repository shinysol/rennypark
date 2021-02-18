using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls.Helper
{
    public static class NumberHelper
    {
        public static int ToInt32(this object obj)
        {
            bool tf = int.TryParse(obj.ToString(), out int result);
            return tf ? result : 0;
        }
        public static long ToInt64(this object obj)
        {
            bool tf = long.TryParse(obj.ToString(), out long result);
            return tf ? result : 0;
        }
        
        public static decimal ToDecimal(this object obj)
        {
            bool tf = decimal.TryParse(obj.ToString(), out decimal result);
            return tf ? result : 0;
        }
        public static float ToFloat(this object obj)
        {
            bool tf = float.TryParse(obj.ToString(), out float result);
            return tf ? result : 0;
        }
        public static double ToDouble(this object obj)
        {
            bool tf = double.TryParse(obj.ToString(), out double result);
            return tf ? result : 0;
        }
    }
}
