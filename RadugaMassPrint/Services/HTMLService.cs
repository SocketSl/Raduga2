using HtmlAgilityPack;
using RadugaMassPrint.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace RadugaMassPrint.Services
{
    internal class HTMLService
    {
        public static void JoinDocuments(IEnumerable<DocumentData> documents, IProgress<(int,int)> progress, CancellationToken token)
        {
            try
            {
                int counter = 0;
                bool isFirstDoc = true;
                StringBuilder contentBuilder = new StringBuilder();
                var htmlDoc = new HtmlDocument();
                foreach (var document in documents.Select(d => d.FileName.Split('/').Last()))
                {
                    token.ThrowIfCancellationRequested();

                    string fullPath = Path.Combine(Environment.CurrentDirectory, document);
                    htmlDoc.Load(fullPath);

                    if (isFirstDoc)
                    {
                        contentBuilder.AppendLine("<html>");
                        var headContent = htmlDoc.DocumentNode.SelectSingleNode("//head");
                        contentBuilder.AppendLine(headContent.InnerHtml);
                        isFirstDoc = false;
                    }
                    var content = htmlDoc.DocumentNode.SelectSingleNode("//body");
                    contentBuilder.AppendLine(content.InnerHtml);
                    contentBuilder.AppendLine("</br>");

                    progress?.Report((2,++counter));
                }

                contentBuilder.AppendLine("</html>");

                using (var streamWriter = new StreamWriter(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "kvit.html")))
                {
                    streamWriter.Write(contentBuilder.ToString());
                }
            }
            catch
            {
                throw;
            }
        }
    }
}
