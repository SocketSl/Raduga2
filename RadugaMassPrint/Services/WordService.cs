using RadugaMassPrint.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

using Word = Microsoft.Office.Interop.Word;

namespace MassPrint.Services
{
    internal static class WordService
    {
        private const string CLEAR_NAME = "clear.doc";
        private const string CLEAR_NAME2 = "clear2.doc";
        internal static async Task PrintDocument(IEnumerable<string> paths, IProgress<(int, int)> progress, CancellationToken token)
        {
            string printerName = "";
            bool accept = false;
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {

                var printDialog = new PrintDialog();

                if (printDialog.ShowDialog() == true)
                {
                    accept = true;
                    printerName = printDialog.PrintQueue.FullName;
                }
            });

            if (accept)
            {
                int counter = 0;

                var wordApp = new Word.Application()
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
                    ActivePrinter = printerName
                };

                try
                {
                    foreach (var path in paths)
                    {
                        token.ThrowIfCancellationRequested(); // Добавить сюда
                        Word.Document document = null;
                        try
                        {
                            document = wordApp.Documents.Open(Path.Combine(Environment.CurrentDirectory, path));
                            document.PrintOut();
                        }
                        catch
                        {
                            document?.Close(false);
                            throw;
                        }
                        progress?.Report((3, ++counter));
                    }
                }
                finally
                {
                    wordApp.Quit();
                }
            }
            else
            {
                foreach (var path in paths)
                {
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                }
            }
        } 

        internal static async Task JoinDocumentsAndPrint(IEnumerable<DocumentData> documents, string folderPath, IProgress<(int, int)> progress, CancellationToken token)
        {
            int counter = 0;

            Word.Application wordApp = null;
            Word.Application wordApp2 = null;

            Word.Document mainDoc = null;
            Word.Document sourceDoc = null;

            try
            {
                wordApp = new Word.Application()
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };

                wordApp2 = new Word.Application()
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };

                mainDoc = wordApp.Documents.Add();


                foreach (var document in documents.Select(d => d.FileName.Split('/').Last()))
                {
                    token.ThrowIfCancellationRequested();
                    mainDoc.Application.Selection.EndKey(Word.WdUnits.wdStory);
                    
                    Word.Range endRange = mainDoc.Content;
                    endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    string fullPath = Path.Combine(Environment.CurrentDirectory, document);

                    try
                    {
                        sourceDoc = wordApp2.Documents.Open(fullPath, Visible: false);
                        sourceDoc.Content.Copy();
                        endRange.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);
                    }
                    finally
                    {
                        if (sourceDoc != null)
                        {
                            sourceDoc.Close(false);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceDoc);
                            sourceDoc = null;
                        }

                        if (endRange != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(endRange);
                            endRange = null;
                        }
                    }
                    progress?.Report((2, ++counter));
                }

                mainDoc.SaveAs2(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), CLEAR_NAME));

                Process.Start(new ProcessStartInfo()
                {
                    FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), CLEAR_NAME),
                    UseShellExecute = true
                });

            }
            catch (Exception e)
            {
                throw;
            }
            finally
            {
                mainDoc?.Close(false);
                wordApp?.Quit();
                wordApp2?.Quit();
                
                if (mainDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mainDoc);
                }

                if (wordApp != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mainDoc);
                }

                if (wordApp2 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mainDoc);
                }

                RemoveAllFiles(documents.Select(d => d.FileName.Split('/').Last()));
            }
        }

        internal static async Task JoinDocumentsWithPageBreakAndPrint(IEnumerable<DocumentData> documents, string folderPath, IProgress<(int, int)> progress, CancellationToken token)
        {
            int counter = 0;
            var docsID = documents.Select(d => d.DocID).ToHashSet().Count();

            Word.Application wordApp = null;
            Word.Application wordApp2 = null;
            Word.Document mainDoc = null;
            Word.Document sourceDoc = null;

            try
            {
                wordApp = new Word.Application()
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
                };


                wordApp2 = new Word.Application()
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
                };

                mainDoc = wordApp2.Documents.Add();

                foreach (var document in documents.Select(d => new { FileName = d.FileName.Split('/').Last(), d.DocID }))
                {
                    token.ThrowIfCancellationRequested();

                    string fullPath = Path.Combine(Environment.CurrentDirectory, document.FileName);
                   
                    Word.Range endRange = mainDoc.Content;
                    endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    if (counter != 0)
                    {
                        endRange.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
                    }

                    Word.Section currentSection = mainDoc.Sections.Last;

                    if (document.DocID == 74)
                    {
                        currentSection.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;

                        // Устанавливаем узкие поля
                        currentSection.PageSetup.LeftMargin = wordApp.InchesToPoints(0.5f);
                        currentSection.PageSetup.RightMargin = wordApp.InchesToPoints(0.5f);
                        currentSection.PageSetup.TopMargin = wordApp.InchesToPoints(0.5f);
                        currentSection.PageSetup.BottomMargin = wordApp.InchesToPoints(0.5f);
                    }

                    else
                    {
                        // Портретная ориентация по умолчанию
                        currentSection.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
                    }

                    // Копируем все содержимое
                    // Открываем исходный документ
                    try
                    {
                        sourceDoc = wordApp.Documents.Open(fullPath, Visible: false);

                        if (document.DocID == 74)
                        {
                            sourceDoc.Content.Font.Size = 9;
                        }
                        sourceDoc.Content.Copy();

                        endRange = mainDoc.Content;
                        endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        endRange.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);

                        if (docsID != 1 && document.DocID == 81)
                        {
                            endRange = mainDoc.Content;
                            endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            endRange.InsertBreak(Word.WdBreakType.wdPageBreak);

                            endRange = mainDoc.Content;
                            endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            endRange.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);
                        }
                    }
                    finally
                    {
                        if (sourceDoc != null)
                        {
                            sourceDoc.Close(false);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceDoc);
                        }
                        if (endRange != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(endRange);
                        }
                    }

                    if (counter < documents.Count() - 1)
                    {
                        endRange = mainDoc.Content;
                        endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        endRange.InsertBreak(Word.WdBreakType.wdPageBreak);
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(currentSection);

                    progress?.Report((2, ++counter));
                }

                string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "PechatnieFormy.doc");
                mainDoc.SaveAs2(outputPath);

                Process.Start(new ProcessStartInfo()
                {
                    FileName = outputPath,
                    UseShellExecute = true
                });

            }
            catch (Exception e)
            {
                throw;
            }
            finally
            {
                mainDoc?.Close(false);
                wordApp?.Quit();
                wordApp2?.Quit();

                if (mainDoc != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mainDoc);
                }
                if (wordApp != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
                if (wordApp2 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp2);
                }  

                RemoveAllFiles(documents.Select(d => d.FileName.Split('/').Last()).ToList());
            }
        }

        private static void RemoveAllFiles(IEnumerable<string> paths)
        {
            foreach (var file in paths)
            {
                if (File.Exists(Path.Combine(Environment.CurrentDirectory, file)))
                {
                    File.Delete(Path.Combine(Environment.CurrentDirectory, file));
                }
            }
        }
    }
}
