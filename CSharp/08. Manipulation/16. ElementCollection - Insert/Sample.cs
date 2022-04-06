using System;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertParagraphCount();
        }
        /// <summary>
        /// Inserts a new Run (Text element) at the start of each paragraph.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/elementcollection-insert.php
        /// </remarks>
        static void InsertParagraphCount()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            int paragraphNum = 1;
            foreach (Element el in dc.Sections[0].GetChildElements(false))
            {
                if (el is Paragraph)
                {
                    // Insert a new Run into Paragraph.InlineCollection 'Inlines'.
                    // InlineCollection is descendant of the base abstract class ElementCollection.
                    (el as Paragraph).Inlines.Insert(0, new Run(dc, "Paragraph " + paragraphNum.ToString() + " - ", new CharacterFormat() { BackgroundColor = Color.Orange, FontColor = Color.White }));
                    paragraphNum++;
                }
            }
            dc.Save("Result.docx");

            // Show the result.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("Result.docx") { UseShellExecute = true });
        }
    }
}