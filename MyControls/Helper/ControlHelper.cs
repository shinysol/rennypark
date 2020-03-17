﻿using System;
using System.Windows.Input;
using System.Windows;
using System.Linq;
using System.Data;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using MyControls.Helper;
using static MyControls.Helper.StringHelper;

namespace MyControls.Helper
{
    public static class ControlHelper
    {
        public static bool IsNullOrWhiteSpace(this TextBox tb)
        {
            if (tb is null) return true;
            return string.IsNullOrWhiteSpace(tb.Text);
        }
        public static string To0or1String(this CheckBox chb)
        {
            try
            {
                return (bool)chb.IsChecked ? "1" : "0";
            }
            catch
            {
                return "-1";
            }
        }
        public static bool IsSelected(this DataGrid dg)
        {
            return !dg.SelectedIndex.Equals(-1);
        }
        public static object SelectedRow(this DataGrid dg, string columnName)
        {
            return (dg.SelectedItem as DataRowView)[columnName];
        }
        public static bool IsSourceNullOrHasZeroRows(this DataGrid dg)
        {
            return dg is null || dg.ItemsSource is null || (dg.ItemsSource as DataView).Table.Rows.Equals(0);
        }
        public static string JoinWithSeparator(this List<string> list, string separator = ", ")
        {
            return string.Join(separator, list.Where(x => !x.IsNullOrWhiteSpace()));
        }
        public static string ForQuery(this TextBox txt)
        {
            return txt.Text.Replace("'", "").Replace(";", "");
        }
        public static void TbxOnlyNumbers(this TextBox tbx)
        {
            tbx.KeyDown += TbxOnlyNumbers;
        }
        private static void TbxOnlyNumbers(object sender, KeyEventArgs e)
        {
            List<int> PermittedKeys = new List<int>();
            for (int i = 34; i <= 43; i++)
            {
                PermittedKeys.Add(i);
            }
            for (int i = 74; i <= 83; i++)
            {
                PermittedKeys.Add(i);
            }
            PermittedKeys.Add((int)Key.OemPeriod);
            PermittedKeys.Add((int)Key.Tab);
            PermittedKeys.Add((int)Key.Decimal);
            if (!PermittedKeys.Contains((int)e.Key))
            {
                e.Handled = true;
            }
        }
    }
}