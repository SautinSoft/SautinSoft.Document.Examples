using System;
using System.Data;
using System.IO;
using System.Collections.Generic;

using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.Globalization;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            TableReportWithRegionsAdvanced();
        }

        /// <summary>
        /// Generates a table report with regions using XML document as a data source and FieldMerging event.
        /// </summary>
        /// <remarks>
        /// See details at: https://www.sautinsoft.com/products/document/help/net/developer-guide/mail-merge-table-report-with-regions-advanced-net-csharp-vb.php
        /// </remarks>
        public static void TableReportWithRegionsAdvanced()
        {
            // Scan directory for image files.
            Dictionary<string, string> icons = new Dictionary<string, string>();
            foreach (string iconPath in Directory.EnumerateFiles(@"..\..\..\icons\", @"*.jpg"))
            {
                icons[Path.GetFileNameWithoutExtension(iconPath.ToLower())] = iconPath;

                switch (Path.GetFileName(iconPath))
                {
                    case "Cherry-apple pie.jpg": icons["Cherry/apple/pie".ToLower()] = iconPath; break;
                    case "Dark-milk-white chocolate.jpg": icons["Dark/milk/white chocolate".ToLower()] = iconPath; break;
                    case "Rice-lemon-vanilla pudding.jpg": icons["Rice/lemon/vanilla pudding".ToLower()] = iconPath; break;
                    case "Spice-cake honey-cake.jpg": icons["Spice-cake, honey-cake".ToLower()] = iconPath; break;
                    case "Chocolate-strawberry-vanilla ice cream.jpg": icons["Chocolate/strawberry/vanilla ice cream".ToLower()] = iconPath; break;
                    default:break;
                }
            }

            // Create the Dataset and read the XML.
            DataSet ds = new DataSet();

            ds.ReadXml(@"..\..\..\Orders.xml");

            // Load the template document.
            string templatePath = @"..\..\..\InvoiceTemplate.docx";

            DocumentCore dc = DocumentCore.Load(templatePath);

            // Each product will be decorated by appropriate icon.
            dc.MailMerge.FieldMerging += (sender, e) =>
            {
                // Insert an icon before the product name
                if (e.RangeName == "Product" && e.FieldName == "Name")
                {
                    e.Inlines.Clear();

                    string iconPath;
                    if (icons.TryGetValue(((string)e.Value).ToLower(), out iconPath))
                    {
                        e.Inlines.Add(new Picture(dc, iconPath) { Layout = new InlineLayout(new Size(30, 30)) });
                        e.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.Tab));
                    }

                    e.Inlines.Add(new Run(dc, (string)e.Value));
                    e.Cancel = false;
                }
                // Add the currency sign into "Total" field.
                // You may change the culture "en-GB" to any desired.
                if (e.RangeName == "Order" && e.FieldName == "OrderTotal")
                {
                    decimal total = 0;
                    if (Decimal.TryParse((string)e.Value, out total))
                    {
                        e.Inlines.Clear();
                        e.Inlines.Add(new Run(dc, String.Format(new CultureInfo("en-GB"), "{0:C}", total)));
                    }
                }
            };

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
