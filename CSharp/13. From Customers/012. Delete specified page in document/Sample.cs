using SautinSoft.Document;
using System;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            DeleteSpecifiedPageInDocument();       
        }
        /// <summary>
        /// How to delete the specified page in the document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-delete-specified-page-in-document-net-csharp-vb.php
        /// </remarks>
        public static void DeleteSpecifiedPageInDocument()
        {
            string inpFile = @"..\..\..\example.docx";
            string outFile = @"result.docx";
            
			// Load a document into DocumentCore.
            DocumentCore dc = DocumentCore.Load(inpFile);

            // Divide the document into separate pages.
            DocumentPaginator dp = dc.GetPaginator();

            // Delete page number two.
            dp.Pages[1].Content.Delete();
            
			 // Save our result as a DOCX file.
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}