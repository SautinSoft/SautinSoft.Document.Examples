using SautinSoft.Document;
using SautinSoft.Document.Tables;
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

            RTLTable();
        }

        /// <summary>
        /// How to add Rigth-to-Left text in a table.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/right-to-left-table.php
        /// </remarks>
        public static void RTLTable()
        {
            string sourcePath = @"..\..\..\RTL.docx";
            string destPath = "RTL.pdf";
            DocumentCore dc = DocumentCore.Load(sourcePath);
            // Show line numbers on the right side of the page
            var pageSetup = dc.Sections[0].PageSetup;
            pageSetup.LineNumberRestartSetting = LineNumberRestartSetting.Continuous;
          
            // Create a new right-to-left paragraph
            var paragraph = new Paragraph(dc);
            paragraph.ParagraphFormat.RightToLeft = true;
            paragraph.Inlines.Add(new Run(dc, "أخذ عن موالية الإمتعاض"));
           
            dc.Sections[0].Blocks.Add(paragraph);

            // Create a right-to-left table with some data inside. 
            var table = new Table(dc);
           
            table.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);
            var row = new TableRow(dc);
            table.TableFormat.Borders.SetBorders(MultipleBorderTypes.Inside | MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1);
            table.Rows.Add(row);

            var firstCellPara = new Paragraph(dc, "של תיבת תרומה מלא");
            firstCellPara.ParagraphFormat.RightToLeft = true;
            row.Cells.Add(new TableCell(dc, firstCellPara));

            var secondCellPara = new Paragraph(dc, "200");
            row.Cells.Add(new TableCell(dc, secondCellPara));
            dc.Sections[0].Blocks.Add(table);

            // Save the document as PDF.
            dc.Save(destPath, new PdfSaveOptions());

            // Show the source and the dest documents.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(sourcePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(destPath) { UseShellExecute = true });
        }
    }
}