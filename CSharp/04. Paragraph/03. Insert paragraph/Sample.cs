using System;
using System.IO;
using System.Linq;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free 100-day key here:   
            // https://sautinsoft.com/start-for-free/

            InsertParagraph();
        }
        /// <summary>
        /// Inserts a new paragraph into an existing PDF document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-paragraphs-to-pdf-document-net-csharp-vb.php
        /// </remarks>
        static void InsertParagraph()
        {
            string inpFile = @"..\..\..\example.pdf";
            string outFile = @"Result.pdf";

            DocumentCore dc = DocumentCore.Load(inpFile);
            Paragraph p = new Paragraph(dc);
            p.Content.Start.Insert("William Shakespeare is an English poet " +
                "and playwright, recognized as the greatest English-language writer, " +
                "the national poet of England and one of the outstanding playwrights of the world.",
                new CharacterFormat() { Size = 20, FontName = "Verdana", FontColor = new Color("#358CCB") });
            p.ParagraphFormat.Alignment = HorizontalAlignment.Justify;

            // Insert the paragraph as 1st element in the 1st section.
            dc.Sections[0].Blocks.Insert(0, p);

            dc.Save(outFile);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });

        }
    }
}