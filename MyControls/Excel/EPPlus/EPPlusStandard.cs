using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls.Excel.EPPlus
{
    public class EPPlusStandard
    {
        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (FileStream stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                using (DataTable tbl = new DataTable())
                {
                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                    }
                    var startRow = hasHeader ? 2 : 1;
                    for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                        DataRow row = tbl.Rows.Add();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                    }
                    return tbl;
                }
            }
        }

        public static void SaveExcelFromDataTable(string path, DataTable dataTable)
        {
            // https://stackoverflow.com/questions/13669733/export-datatable-to-excel-with-epplus
            using (OfficeOpenXml.ExcelPackage pck = new OfficeOpenXml.ExcelPackage(new FileInfo(path)))
            {
                OfficeOpenXml.ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Accounts");
                ws.Cells["A1"].LoadFromDataTable(dataTable, true);
                pck.Save();
            }
        }
    }
}
