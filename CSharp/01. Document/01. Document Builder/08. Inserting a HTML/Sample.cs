using System;
using SautinSoft.Document;
using System.Text;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertingHtml();
        }
        /// <summary>
        /// Inserts an HTML string into the document using DocumentBuilder.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-html.php
        /// </remarks>

        static void InsertingHtml()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            db.InsertHtml(
                "<body>" +
                "<p align='center' style='font-size: 36px'>Welcome to SautinSoft!</p>" +
                 "<p><img src='../../../banner_sautinsoft.jpg' width='400' alt='image from disk'/></p>" +
                 "<p><img src='https://www.sautinsoft.com/images/banner_sautinsoft.jpg' width='200' height='50' alt='image from URL'/></p>" +
                  "<h1 align='center'>Heading 1 center.</h1>" + 
                  "<ol>" +
               "<li>Ordered List 1</li>" + "<ol>" + "<li>Ordered List 1.1</li>" + "<li>Ordered List 1.2</li>" + "</ol>" +
               "<li>Ordered List 2</li>" +
               "</ol>" +

                   "<ul>" +
         "<li>Unordered list 1</li>" +
         "<li>Unordered list 2</li>" +
         "</ul>" +

                "<p align='right'>Paragraph right</p>" +
                "<div align='center'>Div center</div>" +
                "<h1 align='left'>Heading 2 left.</h1>" +
            "</body>"
                );

            // Save our document into DOCX format.
            string resultPath = @"Result.docx";
            dc.Save(resultPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}