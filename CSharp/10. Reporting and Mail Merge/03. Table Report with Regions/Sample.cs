using System;
using System.Data;
using System.IO;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            TableReportWithRegions();
        }

        /// <summary>
        /// Generates a table report with regions using XML document as a data source.
        /// </summary>
        /// <remarks>
        /// See details at: https://www.sautinsoft.com/products/document/help/net/developer-guide/mail-merge-table-report-with-regions-net-csharp-vb.php
        /// </remarks>
        public static void TableReportWithRegions()
        {
            // Create the Dataset and read the XML.
            DataSet ds = new DataSet();

            ds.ReadXml(@"..\..\..\Orders.xml");

            // Load the template document.
            string templatePath = @"..\..\..\InvoiceTemplate.docx";

            DocumentCore dc = DocumentCore.Load(templatePath);

            // Execute the mail merge.
            dc.MailMerge.Execute(ds.Tables["Order"]);

            string resultPath = "Invoices.pdf";

            // Save the output to file
            dc.Save(resultPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}
