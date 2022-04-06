using System.IO;
using System.Linq;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            Manipulation();
        }
        /// <summary>
        /// Manipulation with ElementCollection. Split 1st Paragraph by separate Runs and insert each Run into a new Paragraph.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/elementcollection-manipulation.php
        /// </remarks>
        static void Manipulation()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            string filePathResult = @"Result-file.pdf";
            Section section = dc.Sections[0];
            Paragraph paragraph = section.Blocks[0] as Paragraph;
            for (int i = 1; i < paragraph.Inlines.Count ; )
            {
                Inline inline = paragraph.Inlines[i];
                paragraph.Inlines.RemoveAt(1);
                section.Blocks.Add(new Paragraph(dc, inline));
            }
            dc.Save(filePathResult);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePathResult) { UseShellExecute = true });
        }
    }
}