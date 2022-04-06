using System;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            FindAndReplace();
        }
        /// <summary>
        /// Find an empty paragraphs in document, replace all tables into paragraphs.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/elementcollection-find-replace.php
        /// </remarks>
        static void FindAndReplace()
        {
            string filePath = @"..\..\..\example.docx";
			string result1 = @"ResultEmptyParagraphs.docx";
			string result2 = @"ResultReplacedTables.docx";

            DocumentCore dc = DocumentCore.Load(filePath);
            foreach (Paragraph par in dc.Sections[0].GetChildElements(false,ElementType.Paragraph))
            {
                if ( par.Inlines.Count == 0)
                {
                    par.Inlines.Add(new Run(dc, "<empty paragraph>", new CharacterFormat() { BackgroundColor = Color.Black, FontColor = Color.White }));
                }
            }
            dc.Save(result1);

            for (int i = 0; i < dc.Sections[0].Blocks.Count; i++)
            {
                if (dc.Sections[0].Blocks[i] is Table)
                {
                    dc.Sections[0].Blocks[i] = new Paragraph(dc,new Run(dc, "HERE WAS THE TABLE", new CharacterFormat() { BackgroundColor = Color.Yellow}));
                }
            }
			dc.Save(result2);

            // Show the result.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(result1) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(result2) { UseShellExecute = true });
        }
    }
}