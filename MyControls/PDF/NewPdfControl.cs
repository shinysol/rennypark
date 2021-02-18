using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using static MyControls.Pdfcontrol;
using System.Diagnostics;
using MyControls.Helper;

namespace MyControls.PDF
{
    public class NewPdfControl
    {
        public string FilePath { get; set; }
        private PdfPage PDFPage { get; set; }
        public NewPdfControl()
        {

        }
        ~NewPdfControl()
        {

        }

        public void Save(bool openAfterSave)
        {
            if (FilePath.IsNullOrEmpty())
            if (openAfterSave) Process.Start(FilePath);
        }
        public void Save(string filePath)
        {
            // TODO NewPdfControl 인스턴스 형태로 쓰기 편하게
        }
    }
}
