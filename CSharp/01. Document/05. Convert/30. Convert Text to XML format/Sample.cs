using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            ConvertFromFile();
        }

        /// <summary>
        /// Convert Text to XML (file to file).
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-text-to-xml-in-csharp-vb.php
        /// </remarks>
        static void ConvertFromFile()
        {
            string inpFile = @"..\..\..\example.txt";
            string outFile = @"result.xml";

            DocumentCore dc = DocumentCore.Load(inpFile);

            // Convert non-tabular data and saves it as Xml file.
            dc.Save(outFile, new XmlSaveOptions { ConvertNonTabularDataToSpreadsheet = true });

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}