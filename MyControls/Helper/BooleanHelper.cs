using System;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls.Helper
{
    public static class BooleanHelper
    {
        public static Visibility ToVisiblity(this bool tf, bool collapse = true)
        {
            if (tf)
            {
                return Visibility.Visible;
            }
            else
            {
                return collapse ? Visibility.Collapsed : Visibility.Hidden;
            }
        }
        public static Visibility BoolToVisiblity(bool tf, bool collapse = true)
        {
            if (tf)
            {
                return Visibility.Visible;
            }
            else
            {
                return collapse ? Visibility.Collapsed : Visibility.Hidden;
            }
        }
    }
}
