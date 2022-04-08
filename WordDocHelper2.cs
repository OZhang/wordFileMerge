using Microsoft.Office.Interop.Word;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using PdfSharp.Drawing;

namespace wordFileMerge
{
    internal class WordDocHelper2
    {
        public static string Merge(string[] filesToMerge, string finalOutputFileName, bool insertPageBreaks, Action<string, string> callback)
        {
            var outputPDFs = new List<string>();
            object missing = Type.Missing;
            object pageBreak = WdBreakType.wdSectionBreakNextPage;

            // Create a new Word application
            _Application wordApplication = new Application();

            try
            {
                // Create a new file based on our template
                Document wordDocument = wordApplication.Documents.Add(
                                              ref missing
                                            , ref missing
                                            , ref missing
                                            , ref missing);

                // Make a Word selection object.
                Selection selection = wordApplication.Selection;

                //Count the number of documents to insert;
                int documentCount = filesToMerge.Length;

                wordApplication.Visible = false;
                wordApplication.ScreenUpdating = false;

                foreach (var filename in filesToMerge)
                {
                    // Use the dummy value as a placeholder for optional arguments
                    Document doc = wordApplication.Documents.Open(filename, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing);
                    doc.Activate();

                    var fileInfo = new FileInfo(filename);
                    

                    object outputFileName = filename.Replace(fileInfo.Extension, ".pdf");
                    outputPDFs.Add(outputFileName.ToString());
                    object fileFormat = WdSaveFormat.wdFormatPDF;

                    // Save document into PDF Format
                    doc.SaveAs(ref outputFileName,
                        ref fileFormat, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing);

                    // Close the Word document, but leave the Word application open.
                    // doc has to be cast to type _Document so that it will find the
                    // correct Close method.                
                    object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                    ((_Document)doc).Close(ref saveChanges, ref missing, ref missing);
                    doc = null;
                }

                // Clean up!
                wordDocument = null;
            }
            catch (Exception ex)
            {
                //I didn't include a default error handler so i'm just throwing the error
                return $"合并失败, 异常消息: {ex.Message}";
            }
            finally
            {
                // Finally, Close our Word application
                wordApplication.Quit(ref missing, ref missing, ref missing);
            }

            using (PdfDocument outPdf = new PdfDocument())
            {
                foreach (var pdfFile in outputPDFs)
                {
                    using (PdfDocument one = PdfReader.Open(pdfFile, PdfDocumentOpenMode.Import))
                        CopyPages(one, outPdf);
                }

                AddPageCounter(outPdf);

                outPdf.Save(finalOutputFileName);
                System.Diagnostics.Process.Start(finalOutputFileName.ToString());
            }

            return $"合并完毕";
        }

        private static void AddPageCounter(PdfDocument outPdf)
        {
            // Make a font and a brush to draw the page counter.
            XFont font = new XFont("Verdana", 8);
            XBrush brush = XBrushes.Black;

            // Add the page counter.
            string noPages = outPdf.Pages.Count.ToString();
            for (int i = 0; i < outPdf.Pages.Count; ++i)
            {
                PdfPage page = outPdf.Pages[i];

                // Make a layout rectangle.
                XRect layoutRectangle = new XRect(0/*X*/, page.Height - font.Height * 2/*Y*/, page.Width/*Width*/, font.Height/*Height*/);

                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                {
                    gfx.DrawString(
                        "Page " + (i + 1).ToString() + " of " + noPages,
                        font,
                        brush,
                        layoutRectangle,
                        XStringFormats.Center);
                }
            }
        }

        static void CopyPages(PdfDocument from, PdfDocument to)
        {
            for (int i = 0; i < from.PageCount; i++)
            {
                to.AddPage(from.Pages[i]);
            }
        }

    }
}
