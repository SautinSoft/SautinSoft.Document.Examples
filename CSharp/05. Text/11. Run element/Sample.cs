using System.IO;
using System.Linq;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            CalculateRuns();
        }
        /// <summary>
        /// Loads an existing DOCX document and calculates all 'Run' objects.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/run-element-text-in-docx-document-net-csharp-vb.php
        /// </remarks>
        static void CalculateRuns()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            string filePathResult = @"Result-file.docx";

            foreach (Paragraph par in dc.GetChildElements(true,ElementType.Paragraph))
            {
                int totalRuns = par.GetChildElements(true, ElementType.Run).Count();

                Run r = new Run(dc, "<<This paragraph contains " + totalRuns.ToString() + " Run(s)>>", new CharacterFormat() { BackgroundColor = Color.Yellow, Size = 10, FontColor = Color.Black });
                par.Content.End.Insert(r.Content);
            }
            dc.Save(filePathResult);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePathResult) { UseShellExecute = true });
        }
    }
}