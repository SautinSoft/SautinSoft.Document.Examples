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
        /// Replace all Run elements with Bold formatting to Italic and mark them by yellow.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/manipulation.php
        /// </remarks>
        static void Manipulation()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            string filePathResult = @"Result-file.pdf";

            foreach (Run run in dc.GetChildElements(true, ElementType.Run))
            {
                if (run.CharacterFormat.Bold == true)
                {
                    run.CharacterFormat.Bold = false;
                    run.CharacterFormat.Italic = true;
                    run.CharacterFormat.BackgroundColor = Color.Yellow;
                }
            }
            dc.Save(filePathResult);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePathResult) { UseShellExecute = true });
        }
    }
}