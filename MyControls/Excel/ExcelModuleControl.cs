using System;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using Microsoft.Office.Interop.Excel;

namespace MyControls.Excel
{
    public class ExcelModuleControl : IDisposable
    {
        // https://stackoverflow.com/questions/158706/how-do-i-properly-clean-up-excel-interop-objects - No double dots, No iterations
        bool disposed = false;
        protected Application xlApp;
        protected Workbook xlWorkbook;
        protected Sheets sheets;
        protected Workbooks workbooks1;
        protected Workbooks workbooks2;
        protected Worksheet worksheet1;
        protected Worksheet mergeWorksheet;
        protected Worksheet setValueSheet;
        protected Range range1;
        protected Range range2;
        protected Range setValueRange;
        protected Range mergeRange;
        protected Range offsetRange;
        protected Borders borders;
        protected PageSetup pageSetup;
        protected SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);
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
                GC.Collect();
                GC.WaitForPendingFinalizers();
                if (!(borders is null)) Marshal.ReleaseComObject(borders);
                if (!(pageSetup is null)) Marshal.ReleaseComObject(pageSetup);
                if (!(range1 is null)) Marshal.ReleaseComObject(range1);
                if (!(range2 is null)) Marshal.ReleaseComObject(range2);
                if (!(worksheet1 is null)) Marshal.ReleaseComObject(worksheet1);
                if (!(mergeWorksheet is null)) Marshal.ReleaseComObject(mergeWorksheet);
                if (!(mergeRange is null)) Marshal.ReleaseComObject(mergeRange);
                if (!(setValueSheet is null)) Marshal.ReleaseComObject(setValueSheet);
                if (!(setValueRange is null)) Marshal.ReleaseComObject(setValueRange);
                if (!(offsetRange is null)) Marshal.ReleaseComObject(offsetRange);
                if (!(workbooks1 is null)) Marshal.ReleaseComObject(workbooks1);
                if (!(workbooks2 is null)) Marshal.ReleaseComObject(workbooks2);
                if (!(sheets is null)) Marshal.ReleaseComObject(sheets);
                if (!(xlWorkbook is null)) Marshal.ReleaseComObject(xlWorkbook);
                if (!(xlApp is null)) Marshal.ReleaseComObject(xlApp);
                range1 = null;
                worksheet1 = null;
                workbooks1 = null;
                workbooks2 = null;
                sheets = null;
                xlWorkbook = null;
                xlApp = null;
            }
            // Free any unmanaged objects here.
            //
            disposed = true;
        }
        public ExcelModuleControl()
        {
            // 따져봐야할것 - 오픈시기, 클로즈시기, dispose시기, 모드?읽기쓰기?
            if (!CheckExcel())
            {
                throw new ApplicationException("Microsoft Excel not found.");
            }
            bool CheckExcel()
            {
                Type OfficeType = Type.GetTypeFromProgID("Excel.Application");
                if (OfficeType is null)
                {
                    return false;
                }
                return true;
            }
        }
        public async Task InitializeAsync()
        {
            await Task.Run(new System.Action(() =>
            {
                xlApp = new Application
                {
                    DisplayAlerts = false,
                    ScreenUpdating = false,
                    SheetsInNewWorkbook = 1          // 워크시트 1개짜리 워크북 생성 위한 변수설정
                };
            }));
        }
        public async Task CreateWorkbookAsync()
        {
            await Task.Run(new System.Action(() =>
            {
                workbooks1 = xlApp.Workbooks;
                xlWorkbook = workbooks1.Add();
            }
            ));
            return;
        }
        public async Task InsertNewSheetAsync()
        {
            worksheet1 = xlWorkbook.Sheets[xlWorkbook.Sheets.Count];
            await Task.Run(new System.Action(() =>
            {
                xlWorkbook.Sheets.Add(After: worksheet1);
            }
            ));
            return;
        }
        public virtual async Task ImportDataTableAsync(System.Data.DataTable dt, int sheetOrder = 1, string range = "A1", string worksheetName = "")
        {
            sheets = xlWorkbook.Sheets;
            worksheet1 = sheets.get_Item(sheetOrder);
            offsetRange = worksheet1.Range[range];
            range1 = offsetRange.Offset[1, 0];
            await Task.Run(() =>
            {
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    worksheet1.Cells[offsetRange.Row, i + offsetRange.Column] = dt.Columns[i].ColumnName;
                }
                range1.CopyFromRecordset(ConvertToRecordset(dt));
                if (!worksheetName.Equals(string.Empty))
                {
                    worksheet1.Name = worksheetName;
                }
            });
        }
        private async Task ImportDataTableAsyncForDataSet(System.Data.DataTable dt, int sheetOrder = 1, string range = "A1", string worksheetName = "")
        {
            worksheet1 = sheets.get_Item(sheetOrder);
            offsetRange = worksheet1.Range[range];
            range1 = offsetRange.Offset[1, 0];
            await Task.Run(new System.Action(() =>
            {
                for (var i = 0; i < dt.Columns.Count; i++)
                {
                    worksheet1.Cells[offsetRange.Row, i + offsetRange.Column] = dt.Columns[i].ColumnName;
                }
                range1.CopyFromRecordset(ConvertToRecordset(dt));
                if (!worksheetName.Equals(string.Empty))
                {
                    worksheet1.Name = worksheetName;
                }
            }
            ));
        }
        public async Task ImportDataSetAsync(DataSet ds)    // 얘가 ImportDataTableAsync를 반복호출하면서 COM GC 실패할거같다. 잘 봐바.
        {
            // 먼저 워크시트 개수 확인
            int i = 1;
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                int dtCount = ds.Tables.Count;
                int wsCount = sheets.Count;
                if (!dtCount.Equals(wsCount))
                {
                    if (dtCount > wsCount)
                    {
                        sheets.Add(Count: dtCount - wsCount);
                    }
                }
            }
            ));
            foreach (System.Data.DataTable dt in ds.Tables)
            {
                await ImportDataTableAsyncForDataSet(dt, i++, "A1", dt.TableName);
            }
        }
        public async Task LoadWorkbookAsync(string filePath, bool isReadOnly = true)
        {
            await Task.Run(new System.Action(() =>
            {
                workbooks2 = xlApp.Workbooks;
                xlWorkbook = workbooks2.Open(Filename: filePath, ReadOnly: isReadOnly);
            }
            ));
        }
        public async Task Merge(int sheet, string range)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                mergeWorksheet = sheets.get_Item(sheet);
                mergeRange = mergeWorksheet.Range[range];
                mergeRange.Merge();
            }
            ));
        }
        public async Task MergeDuplicatedColumnData(int sheet, string startCell)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                mergeWorksheet = sheets.get_Item(sheet);
                range2 = mergeWorksheet.Range[startCell];
                int initRow = range2.Row;
                int currRow = initRow;
                int initColumn = range2.Column;
                string initValue = range2.Value2;
                string currValue = initValue;
                while (!string.IsNullOrEmpty(initValue))
                {
                    while (initValue.Equals(currValue))
                    {
                        range2 = mergeWorksheet.Cells[++currRow, initColumn];
                        currValue = range2.Value2;
                    }
                    // get next row value
                    mergeRange = mergeWorksheet.Range[mergeWorksheet.Cells[initRow, initColumn], mergeWorksheet.Cells[currRow - 1, initColumn]];
                    mergeRange.Merge();
                    initRow = currRow;
                    initValue = currValue;
                }
            }
            ));
        }
        public async Task MergeDuplicatedColumnData(int sheet, string startCell, string subStartCell)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                mergeWorksheet = sheets.get_Item(sheet);
                range1 = mergeWorksheet.Range[startCell];
                int initRowA = range1.Row;
                int initRowB = initRowA;
                int currRow = initRowA;
                int initColumnA = range1.Column;
                string initValueA = range1.Value2;
                string currValueA = initValueA;
                range2 = mergeWorksheet.Range[subStartCell];
                int initColumnB = range2.Column;
                string initValueB = range2.Value2;
                string currValueB = initValueB;
                bool mergeA = false;
                bool mergeB = false;
                while (!string.IsNullOrEmpty(initValueA) && !string.IsNullOrEmpty(initValueB))
                {
                    while (!mergeA && !mergeB)
                    {
                        range1 = mergeWorksheet.Cells[++currRow, initColumnA];
                        range2 = mergeWorksheet.Cells[currRow, initColumnB];
                        currValueA = range1.Value2;
                        currValueB = range2.Value2;
                        mergeA = !initValueA.Equals(currValueA);
                        mergeB = !initValueB.Equals(currValueB);
                    }
                    // get next row value
                    if (mergeA)
                    {
                        mergeRange = mergeWorksheet.Range[mergeWorksheet.Cells[initRowA, initColumnA], mergeWorksheet.Cells[currRow - 1, initColumnA]];
                        mergeRange.Merge();
                        initRowA = currRow;
                        initValueA = currValueA;
                        mergeB = true;  // 소속 바뀌면 직원 자동 merge
                    }
                    if (mergeB)
                    {
                        mergeRange = mergeWorksheet.Range[mergeWorksheet.Cells[initRowB, initColumnB], mergeWorksheet.Cells[currRow - 1, initColumnB]];
                        mergeRange.Merge();
                        initRowB = currRow;
                        initValueB = currValueB;
                    }
                    mergeA = false;
                    mergeB = false;
                }
            }
            ));
        }
        public async Task MergeDuplicatedColumnData(int sheet, string startCell, string subStartCell, string amountStartCell, string sumStartCell)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                mergeWorksheet = sheets.get_Item(sheet);
                range1 = mergeWorksheet.Range[startCell];
                int initRowA = range1.Row;
                int initRowB = initRowA;
                int currRow = initRowA;
                int initColumnA = range1.Column;
                string initValueA = range1.Value2;
                string currValueA = initValueA;
                range2 = mergeWorksheet.Range[subStartCell];
                int initColumnB = range2.Column;
                string initValueB = range2.Value2;
                string currValueB = initValueB;
                setValueRange = mergeWorksheet.Range[amountStartCell];
                int initColumnAmount = setValueRange.Column;
                double personalAmountSum = 0;
                offsetRange = mergeWorksheet.Range[sumStartCell];
                int initColumnSum = offsetRange.Column;
                bool mergeA = false;
                bool mergeB = false;
                while (!string.IsNullOrEmpty(initValueA) && !string.IsNullOrEmpty(initValueB))
                {
                    while (!mergeA && !mergeB)
                    {
                        setValueRange = mergeWorksheet.Cells[currRow, initColumnAmount];
                        range1 = mergeWorksheet.Cells[++currRow, initColumnA];
                        range2 = mergeWorksheet.Cells[currRow, initColumnB];
                        currValueA = range1.Value2;
                        currValueB = range2.Value2;
                        personalAmountSum += Convert.ToDouble(setValueRange.Value);
                        mergeA = !initValueA.Equals(currValueA);
                        mergeB = !initValueB.Equals(currValueB);
                    }
                    // get next row value
                    if (mergeA)
                    {
                        mergeRange = mergeWorksheet.Range[mergeWorksheet.Cells[initRowA, initColumnA], mergeWorksheet.Cells[currRow - 1, initColumnA]];
                        mergeRange.Merge();
                        initRowA = currRow;
                        initValueA = currValueA;
                        mergeB = true;  // 소속 바뀌면 직원 자동 merge
                    }
                    if (mergeB)
                    {
                        mergeRange = mergeWorksheet.Range[mergeWorksheet.Cells[initRowB, initColumnB], mergeWorksheet.Cells[currRow - 1, initColumnB]];
                        mergeRange.Merge();
                        mergeRange = mergeWorksheet.Range[mergeWorksheet.Cells[initRowB, initColumnSum], mergeWorksheet.Cells[currRow - 1, initColumnSum]];
                        mergeRange.Merge();
                        mergeWorksheet.Cells[initRowB, initColumnSum].Value = personalAmountSum;
                        initRowB = currRow;
                        initValueB = currValueB;
                        personalAmountSum = 0;
                    }
                    mergeA = false;
                    mergeB = false;
                }
            }
            ));
        }
        public async Task SetValue(int sheet, string range, string content)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                setValueRange.Value2 = content;
            }
            ));
        }
        public async Task SetFontSize(int sheet, string range, int size)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                setValueRange.Font.Size = size;
            }
            ));
        }
        public async Task SetNumberFormat(int sheet, string range)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                setValueRange.NumberFormat = "#,###";
            }
            ));
        }
        public async Task SetTextFormat(int sheet, string range)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                setValueRange.NumberFormat = "@";
            }
            ));
        }
        public async Task SetFormat(int sheet, string range, string format)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                setValueRange.NumberFormat = format;
            }
            ));
        }
        public async Task SetBold(int sheet, string range)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                setValueRange.Font.Bold = true;
            }));
        }
        public async Task SetCellColor(int sheet, string range, System.Drawing.Color color)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                setValueRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
            }));
        }
        public async Task SetUnderlined(int sheet, string range)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                setValueRange.Font.Underline = true;
            }
            ));
        }
        public async Task SetHorizontalAlign(int sheet, string range, XlHAlign xlHAlign)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                setValueRange.HorizontalAlignment = xlHAlign;
            }
            ));
        }
        public async Task AutoWidth(int sheet = 1, string range = "")
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                if (string.IsNullOrEmpty(range))
                {
                    setValueRange = setValueSheet.UsedRange;
                }
                else
                {
                    setValueRange = setValueSheet.Range[range];
                }
                range1 = setValueRange.Columns;
                range1.AutoFit();
            }
            ));
        }
        public async Task SetRowHeight(int sheet, string range, double height)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                range1 = setValueRange.Rows;
                range1.RowHeight = height;
            }
            ));
        }
        public async Task SetColumnWidth(int sheet, string range, double width)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                range1 = setValueRange.Columns;
                range1.ColumnWidth = width;
            }
            ));
        }
        public async Task SetFitToPages(int sheet, bool wide = true, bool tall = false)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                pageSetup = setValueSheet.PageSetup;
                pageSetup.Zoom = false;
                if (wide)
                {
                    pageSetup.FitToPagesWide = 1;
                }
                if (tall)
                {
                    pageSetup.FitToPagesTall = 1;
                }
            }
            ));
        }
        public async Task SetPrintArea(int sheet, string rangeAbsolute)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueSheet.PageSetup.PrintArea = rangeAbsolute;
            }
            ));
        }
        public async Task Replace(int sheet, string range, string what, string replacement, XlLookAt lookAt = XlLookAt.xlPart, XlSearchOrder searchOrder = XlSearchOrder.xlByRows,
            bool matchCase = false, bool matchByte = false)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                setValueRange.Replace(what, replacement, lookAt, searchOrder, matchCase, matchByte);
            }
            ));
        }
        public async Task SetTable(int sheet, string range, bool currentRegion, XlLineStyle insideLineStyle)
        {
            await Task.Run(new System.Action(() =>
            {
                sheets = xlWorkbook.Sheets;
                setValueSheet = sheets.get_Item(sheet);
                setValueRange = setValueSheet.Range[range];
                if (currentRegion)
                {
                    range1 = setValueRange.CurrentRegion;
                    borders = range1.Borders;
                }
                else
                {
                    borders = setValueRange.Borders;
                }
                borders.Color = System.Drawing.Color.Black;
                borders[XlBordersIndex.xlInsideHorizontal].LineStyle = insideLineStyle;
                borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;
                borders[XlBordersIndex.xlInsideVertical].LineStyle = insideLineStyle;
                borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlThin;
            }
            ));
        }
        public void SaveWorkbook(string filePath)
        {
            xlWorkbook.SaveAs(Filename: filePath);
        }
        public async Task SaveWorkbookAsync(string filePath)
        {
            await Task.Run(new System.Action(() =>
            {
                xlWorkbook.SaveAs(Filename: filePath);
            }
            ));
        }
        public void CloseWorkbook()
        {
            xlWorkbook.Close();
        }
        public async Task CloseWorkbookAsync()
        {
            await Task.Run(new System.Action(() =>
            {
                xlWorkbook.Close();
            }
            ));
        }
        public async Task<System.Data.DataTable> ConvertExcelToDataTableAsync(string FileName)
        {
            return await Task.Run(() =>
            {
                // Code from http://www.c-sharpcorner.com/code/788/how-to-convert-excel-to-datatable-in-C-Sharp.aspx
                // when this doesn't work properly(especially - oledb 12.0 provider is not registered on local machine..)
                // https://stackoverflow.com/questions/6649363/microsoft-ace-oledb-12-0-provider-is-not-registered-on-the-local-machine
                // -> install this : https://www.microsoft.com/en-us/download/details.aspx?id=23734
                DataSet ds = new DataSet();
                using (OleDbConnection objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FileName + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;';"))
                {
                    objConn.Open();
                    OleDbCommand cmd = new OleDbCommand();
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
                    int totalSheet = dt.Rows.Count; // No of excel sheets
                    sheetName = dt.Rows[0]["TABLE_NAME"].ToString();
                    cmd.Connection = objConn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = string.Format("SELECT * FROM [{0}]", sheetName);
                    using (OleDbDataAdapter oleda = new OleDbDataAdapter(cmd))
                    {
                        oleda.Fill(ds, "excelData");
                        objConn.Close();
                    }
                }
                return ds.Tables["excelData"];
            });
        }
        public void Print(bool FitWide = false, bool FitTall = false)
        {
            throw new Exception();
            foreach (Worksheet ws in xlWorkbook.Sheets)
            {
                PageSetup pagesetup = ws.PageSetup;
                if (FitWide || FitTall)
                {
                    pagesetup.Zoom = false;
                    pagesetup.FitToPagesWide = FitWide;
                    pagesetup.FitToPagesTall = FitTall;
                }
                ws.PrintOutEx(1, 1, 1, false);
            }
        }
        public void Show()
        {
            xlApp.DisplayAlerts = true;
            xlApp.ScreenUpdating = true;
            xlApp.ShowWindowsInTaskbar = true;
            xlApp.Visible = true;
            xlApp.WindowState = XlWindowState.xlMaximized;
        }
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
        ~ExcelModuleControl()
        {
            Dispose(false);
        }
    }
}