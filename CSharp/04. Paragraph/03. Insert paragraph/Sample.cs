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
            p.Content.Start.Insert("Alexander Pushkin was a great russian romantic poet " +
                "and writer who is considered by a lot of people as the best russian poet and the founder " +
                "of contemporary russian literature.",
                new CharacterFormat() { Size = 20, FontName = "Verdana", FontColor = new Color("#358CCB") });
            p.ParagraphFormat.Alignment = HorizontalAlignment.Justify;

            // Insert the paragraph as 1st element in the 1st section.
            dc.Sections[0].Blocks.Insert(0, p);

            dc.Save(outFile);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });

        }
    }
}