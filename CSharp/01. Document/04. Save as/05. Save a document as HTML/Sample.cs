using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program 
    {        
        static void Main(string[] args)
        {
            SaveToHtmlFile();
            SaveToHtmlStream();
        }

        /// <summary>
        /// Open an existing document and saves it as HTML files (in the Fixed and Flowing modes).
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-html-net-csharp-vb.php
        /// </remarks>
        static void SaveToHtmlFile()
        {
            string inputFile = @"..\..\..\example.docx";

            DocumentCore dc = DocumentCore.Load(inputFile);           

            string fileHtmlFixed = @"Fixed-as-file.html";
            string fileHtmlFlowing = @"Flowing-as-file.html";

            // Save to HTML file: HtmlFixed.
            dc.Save(fileHtmlFixed, new HtmlFixedSaveOptions()
            {
                Version = HtmlVersion.Html5,
                CssExportMode = CssExportMode.Inline
            });

            // Save to HTML file: HtmlFlowing.
            dc.Save(fileHtmlFlowing, new HtmlFlowingSaveOptions()
            {
                Version = HtmlVersion.Html5,
                CssExportMode = CssExportMode.Inline,
                ListExportMode = HtmlListExportMode.ByHtmlTags
            });

            // Open the results for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fileHtmlFixed) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fileHtmlFlowing) { UseShellExecute = true });

        }

        /// <summary>
        /// Creates a new document and saves it as HTML documents (in the Fixed and Flowing modes) using MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-html-net-csharp-vb.php
        /// </remarks>
        static void SaveToHtmlStream()
        {
            // There variables are necessary only for demonstration purposes.
            byte[] fileData = null;
            string fileHtmlFixed = @"Fixed-as-stream.html";
            string fileHtmlFlowing = @"Flowing-as-stream.html";

            // Assume we already have a document 'dc'.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hey Guys and Girls!");

            // Let's save our document to a MemoryStream.
            using (MemoryStream ms = new MemoryStream())
            {
                // HTML Fixed.
                dc.Save(ms, new HtmlFixedSaveOptions());
                fileData = ms.ToArray();

                File.WriteAllBytes(fileHtmlFixed, fileData);

                // Or HTML flowing.
                dc.Save(ms, new HtmlFlowingSaveOptions());
                fileData = ms.ToArray();

                File.WriteAllBytes(fileHtmlFlowing, fileData);
            }
        }
    }
}