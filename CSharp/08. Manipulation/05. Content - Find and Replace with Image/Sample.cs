using System.IO;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.Linq;
using System.Text.RegularExpressions;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            FindTextAndReplaceImage();
        }

        /// <summary>
        /// Find Text and replace it with a Picture using ContentRange.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/find-replace-content-net-csharp-vb.php
        /// </remarks>
        public static void FindTextAndReplaceImage()
        {
            // Path to a loadable document.
            string loadPath = @"..\..\..\Critique_signature.docx";
            string pictPath = @"..\..\..\Smile.png";

            // Load a document intoDocumentCore.
            DocumentCore dc = DocumentCore.Load(loadPath);

            //Find "<signature>" Text and Replace everywhere with the "Smile.png"
            // Please note, Reverse() makes sure that action replace not affects to Find().
            Regex regex = new Regex(@"<signature>", RegexOptions.IgnoreCase);
            Picture picture = new Picture(dc, InlineLayout.Inline(new Size(50, 50)), pictPath);
            foreach (ContentRange item in dc.Content.Find(regex).Reverse())
            {
                item.Replace(picture.Content);
            }

            // Save our document into PDF format.
            string savePath = @"..\..\Replaced_signature.pdf";
            dc.Save(savePath, new PdfSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(loadPath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(savePath) { UseShellExecute = true });
        }
    }
}