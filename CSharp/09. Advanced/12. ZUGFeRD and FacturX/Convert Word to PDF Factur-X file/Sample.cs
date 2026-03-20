using SautinSoft;
using SautinSoft.Document;
using System;
using System.IO;
using System.Reflection;

namespace Sample
{
    class Sample
    {
        /// <summary>
        /// Convert Word to PDF (Factur-X) using C# and .NET.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-word-to-pdfa-factur-x.php
        /// </remarks>
        static void Main(string[] args)
        {
            string inpFile = @"..\..\..\example.docx";
            string xmlInfo = File.ReadAllText(@"..\..\..\Factur-X\Factur.xml");

            string outFile = @"..\..\..\FacturXResult.pdf";

            DocumentCore dc = DocumentCore.Load(inpFile);

            PdfSaveOptions pdfSO = new PdfSaveOptions()
            {
                // Factur-X is at the same time a full readable invoice in a PDF A/3 format,
                // containing all information useful for its treatment, especially in case of discrepancy or absence of automatic matching with orders and / or receptions,
                // and a set of invoice data presented in an XML structured file conformant to EN16931 (syntax CII D16B), complete or not, allowing invoice process automation.
                FacturXXML = xmlInfo
            };

            dc.Save(outFile, pdfSO);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y


            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }

    }
}

