using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.IO;
using System.Diagnostics;
using System.Windows;
using PdfSharp.Pdf.IO;

namespace MyControls.Helper.PDF
{
    public static class PDFHelper
    {
        public static Task MergeBytesToPDF(string filePath, params byte[][] pdfs)
        {
            return Task.Run(() =>
            {
                try
                {
                    using (PdfDocument merged = new PdfDocument())
                    {
                        foreach (byte[] pdf in pdfs)
                        {
                            using (MemoryStream pdfStream = new MemoryStream(pdf))
                            using (PdfDocument pdfDoc = PdfReader.Open(pdfStream, PdfDocumentOpenMode.Import))
                            {
                                for (int i = 0; i < pdfDoc.PageCount; i++)
                                {
                                    PdfPage page = pdfDoc.Pages[i];
                                    merged.Pages.Add(page);
                                }
                            }

                        }
                        merged.Save(filePath);
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            });
        }
    }
}
