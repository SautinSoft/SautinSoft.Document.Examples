using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace Example
{
    class Program
    {
        private static DocumentCore dc;

        static void Main(string[] args)
        {
            CreateTable();
        }

        /// <summary>
        /// Creates a new cell within a table and fills the it in staggered order.
        /// </summary>
        /// <param name="rowIndex">Row index.</param>
        /// <param name="colIndex">Column index.</param>
        /// <returns>New table cell.</returns>
        static TableCell NewCell(int rowIndex, int colIndex)
        {
            TableCell cell = new TableCell(dc);

            cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1);

            if (colIndex % 2 == 1 && rowIndex % 2 == 0 || colIndex % 2 == 0 && rowIndex % 2 == 1)
            {
                cell.CellFormat.BackgroundColor = Color.Black;
            }

            Run run = new Run(dc, string.Format("Row - {0}; Col - {1}", rowIndex, colIndex));
            run.CharacterFormat.FontColor = Color.Auto;

            cell.Blocks.Content.Replace(run.Content);

            return cell;
        }             

        /// <summary>
        /// Creates a new document with a table.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/tables.php
        /// </remarks>
        static void CreateTable()
        {
            dc = new DocumentCore();           

            string filePath = @"Result-file.docx";

            Table table = new Table(dc, 5, 5, NewCell);

            // Place the 'Table' at the start of the 'Document'.
            // By the way, we didn't create a 'Section' in our document.
            // As we're using 'Content' property, a 'Section' will be created automatically if necessary.
            dc.Content.Start.Insert(table.Content);
            dc.Save(filePath);

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}