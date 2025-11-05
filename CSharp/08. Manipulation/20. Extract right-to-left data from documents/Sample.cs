using SautinSoft.Document;
using System;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            ConvertRTLcontent();
        }

        /// <summary>
        /// How to convert documents with Right-To-Left content to HTML.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-documents-with-right-to-left-content-to-html.php
        /// </remarks>
        public static void ConvertRTLcontent()
        {
            string sourcePath = @"..\..\..\RTL.docx";
            string destPath = "RTL.html";
            
            // Load document with arabic, hindi, hebrew content.
            DocumentCore dc = DocumentCore.Load(sourcePath);
           
            // Save the document as HTML.
            dc.Save(destPath, new HtmlFixedSaveOptions());

            // Show the source and the dest documents.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(sourcePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(destPath) { UseShellExecute = true });
        }
    }
}