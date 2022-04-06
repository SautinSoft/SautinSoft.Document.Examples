using System.IO;
using System.Linq;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            ModifyTable();
        }

        /// <summary>
        /// How to modify an existing table in a document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/modify-table.php
        /// </remarks>
        public static void ModifyTable()
        {
            string sourcePath = @"..\..\..\table.docx";
            string destPath = "Table modified.docx";

            // Load a document with a table.
            DocumentCore dc = DocumentCore.Load(sourcePath);

            // Find a first table in the document.
            Table table = (Table)dc.GetChildElements(true, ElementType.Table).First();

            // Set dashed borders and yellow background for all cells.
            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                {
                    TableCell cell = table.Rows[r].Cells[c];
                    cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dashed, Color.Black, 1);
                    cell.CellFormat.BackgroundColor = new Color("#FFCC00");
                }
            }

            // Save the document as DOCX.
            dc.Save(destPath, new DocxSaveOptions());

            // Show the source and the dest documents.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(sourcePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(destPath) { UseShellExecute = true });
        }
    }
}