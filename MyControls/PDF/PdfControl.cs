using System;
using PdfSharp.Drawing;
using System.Reflection;
using System.IO;
using System.Collections.Generic;

namespace MyControls.PDF
{
    public static class PdfControl
    {
        //public bool PrintPDF(string FilePath)
        //{
        //    bool result = false;
        //    CAcroApp mApp;
        //    CAcroPDDoc pdDoc;
        //    CAcroAVDoc avDoc;
        //    string szStr;
        //    string szName;
        //    int iNum = 0;

        //    // Initialize Acrobat by creating App object
        //    // must modify Embed Interop T->F of 'Acrobat'
        //    mApp = new AcroAppClass();

        //    //Show Acrobat
        //    //mApp.Show();

        //    //set AVDoc object
        //    avDoc = new AcroAVDocClass();
        //    //open the PDF
        //    if (avDoc.Open(FilePath, ""))
        //    {
        //        //set the pdDoc object and get some data
        //        pdDoc = (CAcroPDDoc)avDoc.GetPDDoc();
        //        //compose a message
        //        if(avDoc.PrintPages(0, pdDoc.GetNumPages() - 1, 0, 0, 0))
        //        {
        //            result = true;
        //        }
        //        pdDoc.Close();
        //        mApp.CloseAllDocs();
        //        mApp.Exit();
        //    }
        //    return result;
        //}

        public static double GetPtfromCm(double centimeters)
        {
            return centimeters * 0.39370 * 72;
        }
        public static double GetPtfrommm(double milimeters)
        {
            return milimeters / 10 * 0.39370 * 72;
        }
        public static double GetPtfromInch(double inches)
        {
            return inches * 72;
        }

        public static void DrawTextinCellmm(XGraphics graph, XFont font, XPen pen, string str, Enums.ContentsAlignment align, double x1, double y1, double x2, double y2, bool wraptext, double margin = 1)
        {
            double wdth = graph.MeasureString(str, font).Width;
            double hght = graph.MeasureString(str, font).Height;
            if (wraptext)
            {
                int txtCursor = 0;
                int chrsTmp = 0;
                List<int> chrs = new List<int>();
                double cellWdth = GetPtfrommm(x2) - GetPtfrommm(x1) - 2 * GetPtfrommm(margin);
                while (txtCursor < str.Length - 1)
                {
                    while (graph.MeasureString(str.Substring(chrsTmp, txtCursor - chrsTmp + 1), font).Width < cellWdth)
                    {
                        if (txtCursor < str.Length - 1)
                        {
                            txtCursor += 1;
                        }
                        else
                        {
                            break;
                        }
                    }
                    chrsTmp = txtCursor;
                    chrs.Add(txtCursor);
                }
                chrsTmp = 0;
                int lineno = 0;
                foreach (int cur in chrs)
                {
                    graph.DrawString(str.Substring(chrsTmp, cur - chrsTmp + 1), font, XBrushes.Black,
                                new XRect(GetPtfrommm(x1) + GetPtfrommm(margin), (GetPtfrommm(y1) + GetPtfrommm(margin) - GetPtfrommm(0.7) + hght * lineno), x2 - x1 - 2 *
                                GetPtfrommm(margin), hght), XStringFormats.TopLeft);
                    chrsTmp = cur;
                    lineno += 1;
                }
            }
            else
            {
                switch (align)
                {
                    case Enums.ContentsAlignment.Left:

                        graph.DrawString(str, font, XBrushes.Black,
                            new XRect(GetPtfrommm(x1) + GetPtfrommm(margin), (GetPtfrommm(y1) / 2) + (GetPtfrommm(y2) / 2) - (hght / 2), wdth, hght), XStringFormats.TopLeft);
                        break;
                    case Enums.ContentsAlignment.Right:
                        graph.DrawString(str, font, XBrushes.Black,
                            new XRect(GetPtfrommm(x2) - GetPtfrommm(margin) - wdth, (GetPtfrommm(y1) / 2) + (GetPtfrommm(y2) / 2) - (hght / 2), wdth, hght), XStringFormats.TopLeft);
                        break;
                    case Enums.ContentsAlignment.Center:
                        graph.DrawString(str, font, XBrushes.Black,
                            new XRect((GetPtfrommm(x1) / 2) + (GetPtfrommm(x2) / 2) - (wdth / 2), (GetPtfrommm(y1) / 2) + (GetPtfrommm(y2) / 2) - (hght / 2), wdth, hght), XStringFormats.TopLeft);
                        break;
                }
            }
        }
        public static void RowContentsFill(XGraphics graph, XPen pen, XFont font, int Row, double[] PositionX, double[] PositionY, string[] contents, Enums.ContentsAlignment align)
        {
            for (int i = 0; i <= contents.Length - 1; i++)
            {
                DrawTextinCellmm(graph, font, pen, contents[i], align, PositionX[i], PositionY[Row], PositionX[i + 1], PositionY[Row + 1], false);
            }
        }
        public static void ColumnContentsFill(XGraphics graph, XPen pen, XFont font, int Column, double[] PositionX, double[] PositionY, string[] contents, Enums.ContentsAlignment align)
        {
            for (int i = 0; i <= contents.Length - 1; i++)
            {
                DrawTextinCellmm(graph, font, pen, contents[i], align, PositionX[Column], PositionY[i], PositionX[Column + 1], PositionY[i + 1], false);
            }
        }

        public static void DrawLineCm(XGraphics graph, XPen pen, double x1, double y1, double x2, double y2)
        {
            graph.DrawLine(pen, GetPtfromCm(x1), GetPtfromCm(y1), GetPtfromCm(x2), GetPtfromCm(y2));
        }
        public static void DrawLinemm(XGraphics graph, XPen pen, double x1, double y1, double x2, double y2)
        {
            graph.DrawLine(pen, GetPtfrommm(x1), GetPtfrommm(y1), GetPtfrommm(x2), GetPtfrommm(y2));
        }
        public static void DrawHorizontalLinemm(XGraphics graph, XPen pen, double y1, double x1, double x2)
        {
            graph.DrawLine(pen, GetPtfrommm(x1), GetPtfrommm(y1), GetPtfrommm(x2), GetPtfrommm(y1));
        }
        public static void DrawVerticalLinemm(XGraphics graph, XPen pen, double x1, double y1, double y2)
        {
            graph.DrawLine(pen, GetPtfrommm(x1), GetPtfrommm(y1), GetPtfrommm(x1), GetPtfrommm(y2));
        }

        public static void DrawTable(XGraphics graph, XPen pen, double x1, double y1, double x2, double y2, byte rowCnt, byte clmnCnt, bool noDiv1stRow)
        {
            double cellHeight = (y2 - y1) / rowCnt;
            double cellWidth = (x2 - x1) / clmnCnt;
            graph.DrawRectangle(pen, new XRect(x1, y1, x2 - x1, y2 - y1));  // xbrush 추가
            if (rowCnt >= 2)
            {
                for (int i = 1; i <= rowCnt - 1; i++)
                {
                    graph.DrawLine(pen, x1, y1 + i * cellHeight, x2, y1 + i * cellHeight);
                }
            }
            if (clmnCnt >= 2)
            {
                for (int i = 1; i <= clmnCnt - 1; i++)
                {
                    graph.DrawLine(pen, x1 + i * cellWidth, y1 + (noDiv1stRow ? 1 : 0) * cellHeight, x1 + i * cellWidth, y2);
                }
            }
            if (noDiv1stRow)
            {
                graph.DrawRectangle(new XPen(XColor.Empty), new XSolidBrush(XColor.FromArgb(250, 250, 250)), new XRect(x1, y1, x2 - x1, cellHeight));
            }
        }
        public static void DrawTable(XGraphics graph, XPen pen, double x1, double y1, double x2, double y2, byte rowCnt, byte clmnCnt, double firstClmnWdth
            , bool noDiv1stRow)
        {
            double cellHeight = (y2 - y1) / rowCnt;
            double cellWidth;
            graph.DrawRectangle(pen, new XRect(x1, y1, x2 - x1, y2 - y1));  // xbrush 추가
            if (clmnCnt == 1)
            {
                cellWidth = firstClmnWdth;
            }
            else
            {
                cellWidth = (x2 - (x1 + firstClmnWdth)) / (clmnCnt - 1);
            }
            if (rowCnt >= 2)
            {
                for (int i = 1; i <= rowCnt - 1; i++)
                {
                    graph.DrawLine(pen, x1, y1 + i * cellHeight, x2, y1 + i * cellHeight);
                }
            }
            if (clmnCnt >= 2)
            {
                for (int i = 1; i <= clmnCnt - 1; i++)
                {
                    graph.DrawLine(pen, x1 + firstClmnWdth + (i - 1) * cellWidth, y1 + (noDiv1stRow ? 1 : 0) * cellHeight, x1 + firstClmnWdth + (i - 1) * cellWidth, y2);
                }
            }
            if (noDiv1stRow)
            {
                graph.DrawRectangle(new XPen(XColor.Empty), new XSolidBrush(XColor.FromArgb(250, 250, 250)), new XRect(x1, y1, x2 - x1, cellHeight));
            }
        }
        public static void DrawTable(XGraphics graph, XPen pen, double x1, double y1, double x2, double y2, double[] heights, double[] widths)
        {
            graph.DrawRectangle(pen, new XRect(x1, y1, x2 - x1, y2 - y1));  // xbrush 추가
            int rowCnt = heights.Length;
            int clmnCnt = widths.Length;
            if (rowCnt > 1)
            {
                for (int i = 0; i <= rowCnt - 2; i++)
                {
                    graph.DrawLine(pen, x1, y1 + heights[i], x2, y1 + heights[i]);
                }
            }
            if (clmnCnt > 1)
            {
                for (int i = 0; i <= clmnCnt - 2; i++)
                {
                    graph.DrawLine(pen, x1 + widths[i], y1, x1 + widths[i], y2);
                }
            }

        }
        public static void DrawTableCm(XGraphics graph, XPen pen, double x1, double y1, double x2, double y2, double[] heights, double[] widths, bool AbsoluteCoor)
        {
            graph.DrawRectangle(pen, new XRect(GetPtfromCm(x1), GetPtfromCm(y1), GetPtfromCm(x2) - GetPtfromCm(x1), GetPtfromCm(y2) - GetPtfromCm(y1)));  // xbrush 추가
            if (AbsoluteCoor)
            {
                if (heights != null)
                {
                    int rowCnt = heights.Length;
                    for (int i = 0; i <= rowCnt - 1; i++)
                    {
                        graph.DrawLine(pen, GetPtfromCm(x1), GetPtfromCm(heights[i]), GetPtfromCm(x2), GetPtfromCm(heights[i]));
                    }
                }
                if (widths != null)
                {
                    int clmnCnt = widths.Length;
                    for (int i = 0; i <= clmnCnt - 1; i++)
                    {
                        graph.DrawLine(pen, GetPtfromCm(widths[i]), GetPtfromCm(y1), GetPtfromCm(widths[i]), GetPtfromCm(y2));
                    }
                }
            }
            else
            {
                if (heights != null)
                {
                    int rowCnt = heights.Length;
                    for (int i = 0; i <= rowCnt - 1; i++)
                    {
                        graph.DrawLine(pen, GetPtfromCm(x1), GetPtfromCm(y1) + GetPtfromCm(heights[i]), GetPtfromCm(x2), GetPtfromCm(y1) + GetPtfromCm(heights[i]));
                    }
                }
                if (widths != null)
                {
                    int clmnCnt = widths.Length;
                    for (int i = 0; i <= clmnCnt - 1; i++)
                    {
                        graph.DrawLine(pen, GetPtfromCm(x1) + GetPtfromCm(widths[i]), GetPtfromCm(y1), GetPtfromCm(x1) + GetPtfromCm(widths[i]), GetPtfromCm(y2));
                    }
                }
            }
        }
        public static void DrawTablemm(XGraphics graph, XPen pen, double x1, double y1, double x2, double y2, double[] heights, double[] widths, bool AbsoluteCoor)
        {
            graph.DrawRectangle(pen, new XRect(GetPtfrommm(x1), GetPtfrommm(y1), GetPtfrommm(x2) - GetPtfrommm(x1), GetPtfrommm(y2) - GetPtfrommm(y1)));  // xbrush 추가
            if (AbsoluteCoor)
            {
                if (heights != null)
                {
                    int rowCnt = heights.Length;
                    for (int i = 0; i <= rowCnt - 1; i++)
                    {
                        graph.DrawLine(pen, GetPtfrommm(x1), GetPtfrommm(heights[i]), GetPtfrommm(x2), GetPtfrommm(heights[i]));
                    }
                }
                if (widths != null)
                {
                    int clmnCnt = widths.Length;
                    for (int i = 0; i <= clmnCnt - 1; i++)
                    {
                        graph.DrawLine(pen, GetPtfrommm(widths[i]), GetPtfrommm(y1), GetPtfrommm(widths[i]), GetPtfrommm(y2));
                    }
                }
            }
            else
            {
                if (heights != null)
                {
                    int rowCnt = heights.Length;
                    for (int i = 0; i <= rowCnt - 1; i++)
                    {
                        graph.DrawLine(pen, GetPtfrommm(x1), GetPtfrommm(y1) + GetPtfrommm(heights[i]), GetPtfrommm(x2), GetPtfrommm(y1) + GetPtfrommm(heights[i]));
                    }
                }
                if (widths != null)
                {
                    int clmnCnt = widths.Length;
                    for (int i = 0; i <= clmnCnt - 1; i++)
                    {
                        graph.DrawLine(pen, GetPtfrommm(x1) + GetPtfrommm(widths[i]), GetPtfrommm(y1), GetPtfrommm(x1) + GetPtfrommm(widths[i]), GetPtfrommm(y2));
                    }
                }
            }
        }
        public static void NewDrawTablemm(XGraphics graph, XPen pen, double[] heights, double[] widths, bool AbsoluteCoor)
        {
            graph.DrawRectangle(pen, new XRect(GetPtfrommm(widths[0]), GetPtfrommm(heights[0]), GetPtfrommm(widths[widths.Length - 1]) - GetPtfrommm(widths[0]),
                GetPtfrommm(heights[heights.Length - 1]) - GetPtfrommm(heights[0])));  // xbrush 추가
            if (AbsoluteCoor)
            {
                if (heights != null)
                {
                    int rowCnt = heights.Length;
                    for (int i = 0; i <= rowCnt - 1; i++)
                    {
                        graph.DrawLine(pen, GetPtfrommm(widths[0]), GetPtfrommm(heights[i]), GetPtfrommm(widths[widths.Length - 1]), GetPtfrommm(heights[i]));
                    }
                }
                if (widths != null)
                {
                    int clmnCnt = widths.Length;
                    for (int i = 0; i <= clmnCnt - 1; i++)
                    {
                        graph.DrawLine(pen, GetPtfrommm(widths[i]), GetPtfrommm(heights[0]), GetPtfrommm(widths[i]), GetPtfrommm(heights[heights.Length - 1]));
                    }
                }
            }
            else
            {
                if (heights != null)
                {
                    int rowCnt = heights.Length;
                    for (int i = 0; i <= rowCnt - 1; i++)
                    {
                        graph.DrawLine(pen, GetPtfrommm(widths[0]), GetPtfrommm(heights[0]) + GetPtfrommm(heights[i]), GetPtfrommm(widths[widths.Length - 1]),
                            GetPtfrommm(heights[0]) + GetPtfrommm(heights[i]));
                    }
                }
                if (widths != null)
                {
                    int clmnCnt = widths.Length;
                    for (int i = 0; i <= clmnCnt - 1; i++)
                    {
                        graph.DrawLine(pen, GetPtfrommm(widths[0]) + GetPtfrommm(widths[i]), GetPtfrommm(heights[0]), GetPtfrommm(widths[0]) + GetPtfrommm(widths[i]),
                            GetPtfrommm(heights[heights.Length - 1]));
                    }
                }
            }
        }

        static string MigraDocFilenameFromByteArray(byte[] image)
        {
            return "base64:" +
                   Convert.ToBase64String(image);
        }
        static byte[] LoadImage(string name)    // MigraDoc에서 Resource가져오기 위한 메소드
        {
            var assembly = Assembly.GetExecutingAssembly();

            using (Stream stream = assembly.GetManifestResourceStream(name))
            {
                if (stream == null)
                    throw new ArgumentException("No resource with name " + name);

                int count = (int)stream.Length;
                byte[] data = new byte[count];
                stream.Read(data, 0, count);
                return data;
            }
        }
    }
}
