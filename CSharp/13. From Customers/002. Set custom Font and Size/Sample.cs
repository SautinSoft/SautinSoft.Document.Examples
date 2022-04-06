using System;
using System.IO;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            SetCustomFontAndSize();
        }
        /// <summary>
        /// Convert DOCX document to PDF file (set custom font, size and line spacing).
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-word-to-pdf-set-custom-font-and-size-csharp-vb-net.php
        /// </remarks>
        public static void SetCustomFontAndSize()
        {
            // Path to a loadable document.
            string inpFile = @"..\..\..\example.docx";
            string outFile = @"result set custom font.pdf";

            DocumentCore dc = DocumentCore.Load(inpFile);

            string singleFontName = "Times New Roman";
            float singleFontSize = 8.0f;
            float singleLineSpacing = 0.8f;

            dc.DefaultCharacterFormat.FontName = singleFontName;
            dc.DefaultCharacterFormat.Size = singleFontSize;

            foreach (Element element in dc.GetChildElements(true, ElementType.Run, ElementType.Paragraph))
            {
                if (element is Run)
                {
                    (element as Run).CharacterFormat.FontName = singleFontName;
                    (element as Run).CharacterFormat.Size = singleFontSize;
                }
                else if (element is Paragraph)
                {
                    (element as Paragraph).ParagraphFormat.LineSpacing = singleLineSpacing;
                }
            }
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(inpFile) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}
