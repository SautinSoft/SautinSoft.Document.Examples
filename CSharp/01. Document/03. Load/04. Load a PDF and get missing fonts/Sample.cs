using System;
using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {

        static void Main(string[] args)
        {
            LoadPDF();
        }

        /// <summary>
        /// Load a PDF document and get missing fonts.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-pdf-document-get-missing-fonts-net-csharp-vb.php
        /// </remarks>
        static void LoadPDF()
        {            
            string inpFile = @"..\..\..\fonts.pdf";
            string outFile = "Result.docx";

            // Some PDF documents can use rare fonts which isn't installed in your system.
            // During the loading of a PDF document the component tries to find all necessary 
            // fonts installed in your system.
            // In case of the font is missing, you can find its name in the property 'MissingFonts' (after the loading process).
            DocumentCore dc = DocumentCore.Load(inpFile);
            
            if (SautinSoft.Document.FontSettings.MissingFonts.Count > 0)
            {
                Console.WriteLine("Missing Fonts:");
                foreach (string fontFamily in SautinSoft.Document.FontSettings.MissingFonts)
                    Console.WriteLine(fontFamily);
            }

            // Next, knowing missing fonts, you can install these fonts in your system.

            // Also, you can specify an extra folder where component should find fonts.
            SautinSoft.Document.FontSettings.UserFontsDirectory = @"d:\My Fonts";

            // Furthermore, you can add font substitutes, to use alternative fonts.            
            SautinSoft.Document.FontSettings.AddFontSubstitutes("Melvetika", "Segoe UI");
            Console.WriteLine("We\'ve changed Melvetica font to Segoe UI");

            // Load the document again.
            dc = DocumentCore.Load(inpFile);
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
            Console.ReadKey();
        }
    }
}