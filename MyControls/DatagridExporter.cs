using System.Windows;
using System.Windows.Controls;

namespace MyControls
{
    public static class DatagridExporter
    {
        public static void GridSetter(Grid grd)
        {
            string headerName = "엑셀 저장(_S)";
            foreach (UIElement ui in grd.Children)
            {
                if (ui is DataGrid)
                {
                    ContextMenu cma = new ContextMenu();
                    if (!ReferenceEquals((ui as DataGrid).ContextMenu, null))
                    {
                        MenuItem[] mns = new MenuItem[(ui as DataGrid).ContextMenu.Items.Count];
                        (ui as DataGrid).ContextMenu.Items.CopyTo(mns, 0);
                        foreach (MenuItem mi in mns)
                        {
                            (ui as DataGrid).ContextMenu.Items.Remove(mi);
                            cma.Items.Add(mi);
                        }
                        cma.Items.Add(new Separator());
                    }
                    MenuItem mn = new MenuItem()
                    {
                        Header = headerName
                    };
                    mn.Click += DGridToExcel;
                    cma.Items.Add(mn);
                    (ui as DataGrid).ContextMenu = cma;
                }
                else if (ui is Panel)
                {
                    foreach (UIElement uie in (ui as Panel).Children)
                    {
                        if (uie is DataGrid)
                        {
                            ContextMenu cma = new ContextMenu();
                            if (!ReferenceEquals((uie as DataGrid).ContextMenu, null))
                            {
                                MenuItem[] mns = new MenuItem[(uie as DataGrid).ContextMenu.Items.Count];
                                (uie as DataGrid).ContextMenu.Items.CopyTo(mns, 0);
                                foreach (MenuItem mi in mns)
                                {
                                    (uie as DataGrid).ContextMenu.Items.Remove(mi);
                                    cma.Items.Add(mi);
                                }
                                cma.Items.Add(new Separator());
                            }
                            MenuItem mn = new MenuItem()
                            {
                                Header = headerName
                            };
                            mn.Click += DGridToExcel;
                            cma.Items.Add(mn);
                            (uie as DataGrid).ContextMenu = cma;
                        }
                        else if (uie is Panel)
                        {
                            foreach (UIElement uiee in (uie as Panel).Children)
                            {
                                if (uiee is DataGrid)
                                {
                                    ContextMenu cma = new ContextMenu();
                                    if (!ReferenceEquals((uiee as DataGrid).ContextMenu, null))
                                    {
                                        MenuItem[] mns = new MenuItem[(uiee as DataGrid).ContextMenu.Items.Count];
                                        (uiee as DataGrid).ContextMenu.Items.CopyTo(mns, 0);
                                        foreach (MenuItem mi in mns)
                                        {
                                            (uiee as DataGrid).ContextMenu.Items.Remove(mi);
                                            cma.Items.Add(mi);
                                        }
                                    cma.Items.Add(new Separator());
                                    }
                                    MenuItem mn = new MenuItem()
                                    {
                                        Header = headerName
                                    };
                                    mn.Click += DGridToExcel;
                                    cma.Items.Add(mn);
                                    (uiee as DataGrid).ContextMenu = cma;
                                }
                            }
                        }

                    }
                }
            }
        }
        public static void DGridToExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                using (Excelcontrol xl = new Excelcontrol())
                {
                    DataGrid dg = ((((sender as MenuItem).Parent) as ContextMenu).PlacementTarget as DataGrid);
                    xl.ExportToExcel(((System.Data.DataView)(dg.ItemsSource)).ToTable(), xl.GetVisOrder(dg));
                }
            }
            catch
            {

            }
        }
        public static void DgridToExcelContext(object sender, RoutedEventArgs e)
        {
            try
            {
                using (Excelcontrol xl = new Excelcontrol())
                {
                    DataGrid dg = ((((sender as MenuItem).Parent) as ContextMenu).PlacementTarget as DataGrid);
                    bool[] visOrder = new bool[dg.Columns.Count];
                    for (int i = 0; i <= dg.Columns.Count - 1; i++)
                    {
                        if (dg.Columns[i].Visibility.Equals(Visibility.Visible))
                        {
                            visOrder[i] = true;
                        }
                        else
                        {
                            visOrder[i] = false;
                        }
                    }
                    System.Data.DataView dv = (System.Data.DataView)(((((sender as MenuItem).Parent) as ContextMenu).PlacementTarget as DataGrid).ItemsSource);
                    xl.ExportToExcel(dv.ToTable(), visOrder);
                }
            }
            catch
            {

            }
        }
    }
}
