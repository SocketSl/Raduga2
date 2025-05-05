using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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

        internal static async Task JoinDocumentsAndPrint(ICollection<string> paths, string folderPath, IProgress<(int, int)> progress, CancellationToken token)
        {
            int counter = 0;
            var wordApp = new Word.Application()
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
            };

            var mainDoc = wordApp.Documents.Add();

            try
            {

                foreach (var path in paths)
                {
                    token.ThrowIfCancellationRequested(); // Добавить сюда
                    mainDoc.Application.Selection.EndKey(Word.WdUnits.wdStory);
                    mainDoc.Application.Selection.InsertFile(Path.Combine(Environment.CurrentDirectory, path));
                    progress?.Report((2, ++counter));
                }

                mainDoc.SaveAs2(Path.Combine(Environment.CurrentDirectory, CLEAR_NAME));
                mainDoc.Close(false);
                wordApp.Quit();
                await PrintDocument(new List<string>() { Path.Combine(Environment.CurrentDirectory, CLEAR_NAME) }, progress, token);

            }
            catch (Exception e)
            {
                mainDoc.Close(false);
                wordApp.Quit();
                throw;
            }
            finally
            {
                paths.Add(Path.Combine(Environment.CurrentDirectory, CLEAR_NAME));
                RemoveAllFiles(paths);
            }
        }

        internal static async Task JoinDocumentsWithPageBreakAndPrint(ICollection<string> paths, string folderPath, IProgress<(int, int)> progress, CancellationToken token)
        {
            int counter = 0;
            var wordApp = new Word.Application()
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
            };

            var mainDoc = wordApp.Documents.Add();

            try
            {

                bool first = true;

                foreach (var path in paths)
                {
                    token.ThrowIfCancellationRequested(); // Добавить сюда
                    if (!first)
                    {
                        // Вставка разрыва страницы перед следующим документом
                        object breakType = Word.WdBreakType.wdPageBreak;
                        mainDoc.Application.Selection.InsertBreak(ref breakType);
                    }

                    mainDoc.Application.Selection.EndKey(Word.WdUnits.wdStory);
                    mainDoc.Application.Selection.InsertFile(Path.Combine(Environment.CurrentDirectory, path));
                    first = false;
                    progress?.Report((2, ++counter));
                }

                mainDoc.SaveAs2(Path.Combine(Environment.CurrentDirectory, CLEAR_NAME2));
                mainDoc.Close(false);
                wordApp.Quit();
                await PrintDocument(new List<string>() { Path.Combine(Environment.CurrentDirectory, CLEAR_NAME2) }, progress, token);
            }
            catch (Exception e)
            {
                mainDoc.Close(false);
                wordApp.Quit();
                throw;
            }
            finally
            {
                paths.Add(Path.Combine(Environment.CurrentDirectory, CLEAR_NAME2));
                RemoveAllFiles(paths);
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
