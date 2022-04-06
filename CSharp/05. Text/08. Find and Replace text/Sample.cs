using System;
using System.IO;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            FindAndReplace();
        }
        /// <summary>
        /// Find and replace a specific text in an existing DOCX document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/find-and-replace-text-in-docx-document-net-csharp-vb.php
        /// </remarks>
        public static void FindAndReplace()
        {
            // Path to the loadable document.
            string loadPath = @"..\..\..\example.docx";

            DocumentCore dc = DocumentCore.Load(loadPath);

            // Find "Bean" and Replace everywhere on "Joker"
            Regex regex = new Regex(@"Bean", RegexOptions.IgnoreCase);

            // Start:

            // Please note, Reverse() makes sure that action Replace() doesn't affect to Find().
            foreach (ContentRange item in dc.Content.Find(regex).Reverse())
            {
                item.Replace("Joker", new CharacterFormat() { BackgroundColor = Color.Yellow, FontName = "Arial", Size = 16.0 });
            }

            // End:

            // The code above finds and replaces the content in the whole document.
            // Let us say, you want to replace a text inside shape blocks only:

            // 1. Comment the code above from the line "Start" to the "End".
            // 2. Uncomment this code:
            //foreach (Shape shp in dc.GetChildElements(true, ElementType.Shape).Reverse())
            //{
            //    foreach (ContentRange item in shp.Content.Find(regex).Reverse())
            //    {
            //        item.Replace("Joker", new CharacterFormat() { BackgroundColor = Color.Yellow, FontName = "Arial", Size = 16.0 });
            //    }
            //}

            // Save the document as DOCX format.
            string savePath = Path.ChangeExtension(loadPath, ".replaced.docx");
            dc.Save(savePath, SaveOptions.DocxDefault);

            // Open the original and result documents for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(loadPath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(savePath) { UseShellExecute = true });
        }
    }
}
