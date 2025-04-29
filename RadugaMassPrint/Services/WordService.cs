using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace MassPrint.Services
{
    internal static class WordService
    {
        internal static async Task PrintDocument(IEnumerable<string> paths, IProgress<(int, int)> progress)
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
                        var document = wordApp.Documents.Open(Path.Combine(Environment.CurrentDirectory, path));
                        document.PrintOut();
                        document.Close(false);
                        progress?.Report((2, ++counter));
                    }

                }
                finally
                {
                    wordApp.Quit();
                }
            }
        } 

        internal static void JoinDocuments(IEnumerable<string> paths, string folderPath, IProgress<(int, int)> progress)
        {
            int counter = 0;
            var wordApp = new Word.Application()
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
            };
            var mainDoc = wordApp.Documents.Open(folderPath);

            try
            {

                foreach (var path in paths)
                {
                    mainDoc.Application.Selection.EndKey(Word.WdUnits.wdStory);
                    mainDoc.Application.Selection.InsertFile(Path.Combine(Environment.CurrentDirectory, path));
                    progress?.Report((2, ++counter));
                }

                mainDoc.Save();

            }
            finally
            {
                mainDoc.Close();
                wordApp.Quit();
                RemoveAllFiles(paths);
            }
        }

        internal static void JoinDocumentsWithPageBreak(IEnumerable<string> paths, string folderPath)
        {
            var wordApp = new Word.Application()
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
            };

            var mainDoc = wordApp.Documents.Open(folderPath);

            try
            {

                bool first = true;

                foreach (var path in paths)
                {
                    if (!first)
                    {
                        // Вставка разрыва страницы перед следующим документом
                        object breakType = Word.WdBreakType.wdPageBreak;
                        mainDoc.Application.Selection.InsertBreak(ref breakType);
                    }

                    mainDoc.Application.Selection.EndKey(Word.WdUnits.wdStory);
                    mainDoc.Application.Selection.InsertFile(Path.Combine(Environment.CurrentDirectory, path));
                    first = false;
                }

                mainDoc.Save();
            }
            finally
            {
                mainDoc.Close();
                wordApp.Quit();
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
