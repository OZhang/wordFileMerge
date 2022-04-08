using Microsoft.Office.Interop.Word;
using System;
using System.IO;

namespace wordFileMerge
{
    internal class WordDocHelper
    {
        public static string Merge(string[] filesToMerge, string outputFilename, bool insertPageBreaks, Action<string, string> callback)
        {
            object missing = Type.Missing;
            object pageBreak = WdBreakType.wdSectionBreakNextPage;
            object outputFile = outputFilename;

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

                //A counter that signals that we shoudn't insert a page break at the end of document.
                int breakStop = 0;

                // Loop thru each of the Word documents
                foreach (string file in filesToMerge)
                {
                    var fileInfo = new FileInfo(file);
                    breakStop++;
                    callback(fileInfo.Name, $"{breakStop} / {documentCount}");
                    // Insert the files to our template
                    selection.InsertFile(
                                                file
                                            , ref missing
                                            , ref missing
                                            , ref missing
                                            , ref missing);

                    //Do we want page breaks added after each documents?
                    if (insertPageBreaks && breakStop != documentCount)
                    {
                        selection.InsertBreak(ref pageBreak);
                    }
                }

                // If the page count isn't even, add a blank page
                var numberOfPages = wordDocument.ComputeStatistics(WdStatistic.wdStatisticPages, false);
                if (numberOfPages % 2 != 0)
                {
                    selection.InsertBreak(ref pageBreak);
                }

                AddPageNum(missing, wordDocument);

                // Save the document to it's output file.
                wordDocument.SaveAs(
                                ref outputFile
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing);

                System.Diagnostics.Process.Start(outputFile.ToString());
                // Clean up!
                wordDocument = null;
                return $"合并完毕, 页数: {numberOfPages}";
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
        }

        private static object AddPageNum(object missing, Document wordDocument)
        {
            wordDocument.ActiveWindow.View.Type = WdViewType.wdPrintView;
            wordDocument.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekPrimaryFooter; ;
            wordDocument.ActiveWindow.Selection.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordDocument.ActiveWindow.Selection.Font.Name = "Arial";
            wordDocument.ActiveWindow.Selection.Font.Size = 8;
            Object CurrentPage = WdFieldType.wdFieldPage;
            wordDocument.ActiveWindow.Selection.Fields.Add(wordDocument.ActiveWindow.Selection.Range, ref CurrentPage, ref missing, ref missing);
            return missing;
        }
    }
}
