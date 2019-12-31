using System;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls.Control
{
    public static class TcmsControl
    {
        public static string ControlTextExtractor(System.Windows.Controls.Control control, bool comboboxIndex = false)
        {
            if (control is TextBox)
            {
                return (control as TextBox).Text;
            }
            else if (control is ComboBox)
            {
                if (comboboxIndex)
                {
                    return (control as ComboBox).SelectedIndex.ToString();
                }
                else
                {
                    return (control as ComboBox).Text;
                }
            }
            else if (control is DatePicker)
            {
                return ((DateTime)((control as DatePicker).SelectedDate)).ToShortDateString();
            }
            else if (control is CheckBox)
            {
                return ((bool)(control as CheckBox).IsChecked) ? "1" : "0";
            }
            else if (control is Label)
            {
                return (control as Label).Content.ToString();
            }
            else
            {
                throw new NotImplementedException();
            }
        }
    }
}
