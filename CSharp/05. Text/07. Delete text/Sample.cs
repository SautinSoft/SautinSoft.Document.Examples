using System;
using System.Linq;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            DeleteText();
        }
        /// <summary>
        /// Delete a specific text from DOCX document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/delete-text-from-docx-document-net-csharp-vb.php
        /// </remarks>
        static void DeleteText()
        {
            string filePath = @"..\..\..\example.docx";
            string fileResult = @"Result.pdf";
            string textToDelete = "document";
            DocumentCore dc = DocumentCore.Load(filePath);

            int countDel = 0;
            foreach (ContentRange cr in dc.Content.Find(textToDelete).Reverse())
            {
                cr.Delete();
                countDel++;
            }
            Console.WriteLine("The text: \"" + textToDelete + "\" - was deleted " + countDel.ToString() + " time(s).");
            Console.ReadKey();

            dc.Save(fileResult);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fileResult) { UseShellExecute = true });

        }
    }
}