using System;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using Microsoft.Office.Interop.Excel;

namespace MyControls
{
    public class Excelcontrol : IDisposable
    {
        // string resources 1.0.11152.1 현재(18/6/25) 미반영
        public const string INITIALIZE_EXCEL = "엑셀 실행 중";
        public const string CREATE_WORKBOOK = "워크북 생성 중";
        public const string PASTING_DATA = "데이터 붙여넣는 중";
        public const string ADJUSTING_COLUMNWIDTH = "열 너비 정리 중";
        public const string DONE_EXTRACTING = "추출 완료";
        bool disposed = false;
        SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                handle.Dispose();
                // Free any other managed objects here.
                //
            }

            // Free any unmanaged objects here.
            //
            disposed = true;
        }
        ~Excelcontrol()
        {
            Dispose(false);
        }
        // Export
        protected bool CheckExcel()
        {
            Type OfficeType = Type.GetTypeFromProgID("Excel.Application");
            if (OfficeType == null)
            {
                return false;
            }
            return true;
        }
        // Following codes are from https://www.codeproject.com/Articles/10503/Simplest-code-to-convert-an-ADO-NET-DataTable-to-a
        static public ADODB.Recordset ConvertToRecordset(System.Data.DataTable inTable)
        {
            ADODB.Recordset result = new ADODB.Recordset()
            {
                CursorLocation = ADODB.CursorLocationEnum.adUseClient
            };
            ADODB.Fields resultFields = result.Fields;
            System.Data.DataColumnCollection inColumns = inTable.Columns;

            foreach (DataColumn inColumn in inColumns)
            {
                resultFields.Append(inColumn.ColumnName
                    , TranslateType(inColumn.DataType)
                    , inColumn.MaxLength
                    , inColumn.AllowDBNull ? ADODB.FieldAttributeEnum.adFldIsNullable :
                                             ADODB.FieldAttributeEnum.adFldUnspecified
                    , null);
            }

            result.Open(System.Reflection.Missing.Value
                    , System.Reflection.Missing.Value
                    , ADODB.CursorTypeEnum.adOpenStatic
                    , ADODB.LockTypeEnum.adLockOptimistic, 0);

            foreach (DataRow dr in inTable.Rows)
            {
                result.AddNew(System.Reflection.Missing.Value,
                              System.Reflection.Missing.Value);

                for (int columnIndex = 0; columnIndex < inColumns.Count; columnIndex++)
                {
                    resultFields[columnIndex].Value = dr[columnIndex];
                }
            }

            return result;
        }
        static ADODB.DataTypeEnum TranslateType(Type columnType)
        {
            switch (columnType.UnderlyingSystemType.ToString())
            {
                case "System.Boolean":
                    return ADODB.DataTypeEnum.adBoolean;

                case "System.Byte":
                    return ADODB.DataTypeEnum.adUnsignedTinyInt;

                case "System.Char":
                    return ADODB.DataTypeEnum.adChar;

                case "System.DateTime":
                    return ADODB.DataTypeEnum.adDate;

                case "System.Decimal":
                    return ADODB.DataTypeEnum.adCurrency;

                case "System.Double":
                    return ADODB.DataTypeEnum.adDouble;

                case "System.Int16":
                    return ADODB.DataTypeEnum.adSmallInt;

                case "System.Int32":
                    return ADODB.DataTypeEnum.adInteger;

                case "System.Int64":
                    return ADODB.DataTypeEnum.adBigInt;

                case "System.SByte":
                    return ADODB.DataTypeEnum.adTinyInt;

                case "System.Single":
                    return ADODB.DataTypeEnum.adSingle;

                case "System.UInt16":
                    return ADODB.DataTypeEnum.adUnsignedSmallInt;

                case "System.UInt32":
                    return ADODB.DataTypeEnum.adUnsignedInt;

                case "System.UInt64":
                    return ADODB.DataTypeEnum.adUnsignedBigInt;

                case "System.String":
                default:
                    return ADODB.DataTypeEnum.adVarChar;
            }
        }
        public bool Print(string FilePath, bool FitWide = false, bool FitTall = false)
        {
            Application xlApp = new Application();
            if (!CheckExcel())
            {
                return false;
            }
            try
            {
                xlApp.DisplayAlerts = false;
                Workbook xlWorkbook;
                xlWorkbook = xlApp.Workbooks.Open(FilePath);
                foreach(Worksheet ws in xlWorkbook.Sheets)
                {
                    if(FitWide || FitTall)
                    {
                        ws.PageSetup.Zoom = false;
                        if (FitWide)
                        {
                            ws.PageSetup.FitToPagesWide = true;
                        }
                        if (FitTall)
                        {
                            ws.PageSetup.FitToPagesTall = true;
                        }
                    }
                    ws.PrintOutEx(1, 1, 1, false);
                    Marshal.ReleaseComObject(ws);
                }
                xlWorkbook.Close(false);
                xlApp.DisplayAlerts = true;
                xlApp.ScreenUpdating = true;
                xlApp.Quit();
                while (Marshal.ReleaseComObject(xlApp) != 0) { }
                while (Marshal.ReleaseComObject(xlWorkbook) != 0) { }
                xlApp = null;
                xlWorkbook = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                return true;
            }
            catch
            {
                return false;
            }
        }
        public bool ExportToExcel(System.Data.DataTable dt)
        {
            //bool[] dc;
            Application xlApp = new Application();
            if (!CheckExcel())
            {
                return false;
            }
            try
            {
                xlApp.DisplayAlerts = false;
                Workbook xlWorkbook;
                Worksheet xlWorkSheet;
                Worksheet ws_2b_del;
                xlWorkbook = xlApp.Workbooks.Add();
                xlWorkSheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);
                //시트 정리작업(1개로)
                int Wsht_Cnt;
                Wsht_Cnt = xlWorkbook.Worksheets.Count;
                if (Wsht_Cnt >= 2)
                {
                    for (var i = 2; i <= Wsht_Cnt; i++)
                    {
                        ws_2b_del = xlWorkbook.Worksheets[2];
                        ws_2b_del.Delete();
                    }
                }
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    xlWorkSheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                }
                xlWorkSheet.Range["A2"].CopyFromRecordset(ConvertToRecordset(dt));
                xlApp.DisplayAlerts = true;
                xlApp.ShowWindowsInTaskbar = true;
                xlApp.Visible = true;
                xlApp.WindowState = XlWindowState.xlMaximized;
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);
                return true;
            }
            catch
            {
                return false;
            }
        }
        public bool ExportToExcel(System.Data.DataTable dt, bool[] visOrder)
        {
            Application xlApp = new Application();
            if (!CheckExcel())
            {
                return false;
            }
            try
            {
                xlApp.DisplayAlerts = false;
                Workbook xlWorkbook;
                Worksheet xlWorkSheet;
                Worksheet ws_2b_del;
                xlWorkbook = xlApp.Workbooks.Add();
                xlWorkSheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);
                //시트 정리작업(1개로)
                int Wsht_Cnt;
                Wsht_Cnt = xlWorkbook.Worksheets.Count;
                if (Wsht_Cnt >= 2)
                {
                    for (var i = 2; i <= Wsht_Cnt; i++)
                    {
                        ws_2b_del = xlWorkbook.Worksheets[2];
                        ws_2b_del.Delete();
                    }
                }
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    xlWorkSheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                }
                xlWorkSheet.Range["A2"].CopyFromRecordset(ConvertToRecordset(dt));
                for (int i=visOrder.Length - 1; i>=0; i--)
                {
                    if (!visOrder[i])
                    {
                        xlWorkSheet.Columns[i + 1].Delete();
                    }
                }
                xlApp.DisplayAlerts = true;
                xlApp.ShowWindowsInTaskbar = true;
                xlApp.Visible = true;
                xlApp.WindowState = XlWindowState.xlMaximized;
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);
                return true;
            }
            catch
            {
                return false;
            }
        }
        public bool[] GetVisOrder(System.Windows.Controls.DataGrid dg)
        {
            bool[] visOrder = new bool[dg.Columns.Count];
            for (int i = 0; i <= dg.Columns.Count - 1; i++)
            {
                if (dg.Columns[i].Visibility.Equals(System.Windows.Visibility.Visible))
                {
                    visOrder[i] = true;
                }
                else
                {
                    visOrder[i] = false;
                }
            }
            return visOrder;
        }
        public virtual bool ExportToExcel(System.Data.DataTable dt, string filepath)
        {
            Application xlApp = new Application();
            if (!CheckExcel())
            {
                return false;
            }
            try
            {
                xlApp.DisplayAlerts = false;
                xlApp.ScreenUpdating = false;
                Workbook xlWorkbook;
                Worksheet xlWorkSheet;
                Worksheet ws_2b_del;
                xlWorkbook = xlApp.Workbooks.Add();
                xlWorkbook.SaveAs(Filename: filepath);
                xlWorkSheet = (Worksheet)xlWorkbook.Worksheets.get_Item(1);
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    xlWorkSheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                }
                xlWorkSheet.Range["A2"].CopyFromRecordset(ConvertToRecordset(dt));
                //시트 정리작업(1개로)
                int Wsht_Cnt;
                Wsht_Cnt = xlWorkbook.Worksheets.Count;
                if (Wsht_Cnt >= 2)
                {
                    for (var i = 2; i <= Wsht_Cnt; i++)
                    {
                        ws_2b_del = xlWorkbook.Worksheets[2];
                        ws_2b_del.Delete();
                    }
                }
                xlWorkSheet.Name = "Export";
                xlWorkbook.SaveAs(Filename: filepath);
                xlWorkbook.Close(SaveChanges: false);
                xlApp.DisplayAlerts = true;
                xlApp.ScreenUpdating = true;
                xlApp.Quit();
                // Manual disposal because of COM
                while (Marshal.ReleaseComObject(xlApp) != 0) { }
                while (Marshal.ReleaseComObject(xlWorkbook) != 0) { }
                xlApp = null;
                xlWorkbook = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                return true;
            }
            catch
            {
                return false;
            }
        }
        public System.Data.DataTable ConvertExcelToDataTable(string FileName)
        {
            // Code from http://www.c-sharpcorner.com/code/788/how-to-convert-excel-to-datatable-in-C-Sharp.aspx
            // when this doesn't work properly(especially - oledb 12.0 provider is not registered on local machine..)
            // https://stackoverflow.com/questions/6649363/microsoft-ace-oledb-12-0-provider-is-not-registered-on-the-local-machine
            // -> install this : https://www.microsoft.com/en-us/download/details.aspx?id=23734
            System.Data.DataTable dtResult = null;
            int totalSheet = 0; //No of sheets on excel file  
            using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
            {
                objConn.Open();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                DataSet ds = new DataSet();
                System.Data.DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                if (dt != null)
                {
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                         where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                         select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                }
                cmd.Connection = objConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = string.Format("SELECT * FROM [{0}]",sheetName);
                oleda = new OleDbDataAdapter(cmd);
                oleda.Fill(ds, "excelData");
                dtResult = ds.Tables["excelData"];
                objConn.Close();
                return dtResult; //Returning Dattable  
            }
        }
        public async Task<System.Data.DataTable> ConvertExcelToDataTableAsync(string FileName)
        {
            return await Task.Run(new Func<System.Data.DataTable>(() =>
            {
                // Code from http://www.c-sharpcorner.com/code/788/how-to-convert-excel-to-datatable-in-C-Sharp.aspx
                int totalSheet = 0; //No of sheets on excel file  
                DataSet ds = new DataSet();
                using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';"))
                {
                    objConn.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    OleDbDataAdapter oleda = new OleDbDataAdapter();
                    System.Data.DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sheetName = string.Empty;
                    if (dt is null)
                    {
                        return null;
                    }
                    var tempDataTable = (from dataRow in dt.AsEnumerable()
                                            where !dataRow["TABLE_NAME"].ToString().Contains("FilterDatabase")
                                            select dataRow).CopyToDataTable();
                    dt = tempDataTable;
                    totalSheet = dt.Rows.Count;
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                    cmd.Connection = objConn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = string.Format("SELECT * FROM [{0}]", sheetName);
                    oleda = new OleDbDataAdapter(cmd);
                    oleda.Fill(ds, "excelData");
                    objConn.Close();
                }
                return ds.Tables["excelData"];
            }));
        }
    }
}
