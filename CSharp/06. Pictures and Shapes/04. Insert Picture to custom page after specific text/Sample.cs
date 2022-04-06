using System.Linq;
using System.IO;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertPictureToCustomPageAndPosition();
        }
        /// <summary>
        /// Insert a picture to custom page and position into existing DOCX document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-picture-jpg-image-to-custom-docx-page-and-position-net-csharp-vb.php
        /// </remarks>
        static void InsertPictureToCustomPageAndPosition()
        {
            // In this example we'll insert the picture to 1st after the word "Sign:".            

            string inpFile = @"..\..\..\example.docx";
            string outFile = @"Result.docx";
            string pictFile = @"..\..\..\picture.jpg";

            DocumentCore dc = DocumentCore.Load(inpFile);
            DocumentPaginator dp = dc.GetPaginator();

            // Find the text "Sign:" on the 1st page.
            ContentRange cr = dp.Pages[0].Content.Find("Sign:").LastOrDefault();
            if (cr != null)
            {
                Picture pic = new Picture(dc, pictFile);
                cr.End.Insert(pic.Content);
            }
            // Save the document as new DOCX and open it.
            dc.Save(outFile);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}