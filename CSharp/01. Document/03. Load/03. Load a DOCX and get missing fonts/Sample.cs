using System;
using System.Collections.Generic;
using System.Linq;
using SautinSoft.Document;

namespace Example
{
    class Program
    {

        static void Main(string[] args)
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            LoadPDF();
        }

        /// <summary>
        /// Load a DOCX document, get and replace missing fonts.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-pdf-document-get-missing-fonts-net-csharp-vb.php
        /// </remarks>
        static void LoadPDF()
        {
            string inpFile = @"..\..\..\fonts.docx";
            string outFile = "Result.pdf";

            // Some documents can use embedded or rare fonts which isn't installed in your system.
            // This triggers the FontSelection event, where, for example, you can output
            // a list of lost fonts to the console or add a replacement for them.
            DocumentCore dc = DocumentCore.Load(inpFile);

            List<string> MissingFonts = new List<string>();
            FontSettings.FontSelection += (s, e) =>
            {
                // Search for embedded and missing fonts.
                if (!MissingFonts.Contains(e.FontName))
                    MissingFonts.Add(e.FontName);

                // Replacing the Dubai font with Segoe UI.
                var segoe = FontSettings.Fonts.FirstOrDefault(f => f.FamilyName == "Segoe UI");
                if (e.FontName == "Dubai")
                {
                    e.SelectedFont = segoe;
                    Console.WriteLine("We\'ve changed Dubai font to Segoe UI");
                }
            };

            dc.Save(outFile);

            if (MissingFonts.Count > 0)
            {
                Console.WriteLine("Missing Fonts:");
                foreach (string fontFamily in MissingFonts)
                    Console.WriteLine(fontFamily);
            }

            // Next, knowing missing fonts, you can install these fonts in your system.

            // Also, you can specify an extra folder where component should find fonts.
            FontSettings.FontsBaseDirectory = @"d:\My Fonts";

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}