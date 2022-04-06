using System;
using SautinSoft.Document;
using SautinSoft.Document.Tables;


namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            DetectBorders();
        }
        /// <summary>
        /// Detect cell borders with the same color.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-detect-borders-in-table-csharp-vb-net.php
        /// </remarks>

        private static void DetectBorders()
            {
                DocumentCore dc = DocumentCore.Load(@"..\..\..\example.docx");

            foreach (TableCell itemTC in dc.GetChildElements(true, ElementType.TableCell))
            {
                SingleBorder sbLeft = itemTC.CellFormat.Borders[SingleBorderType.Left];
                SingleBorder sbTop = itemTC.CellFormat.Borders[SingleBorderType.Top];
                SingleBorder sbRight = itemTC.CellFormat.Borders[SingleBorderType.Right];
                SingleBorder sbBottom = itemTC.CellFormat.Borders[SingleBorderType.Bottom];
                if (sbLeft.Color == sbTop.Color && sbTop.Color == sbRight.Color && sbRight.Color == sbBottom.Color)
                {
                    itemTC.Content.Start.Insert("This cell has the same border color.\r\n");
                }
            }

            // Save our document into DOCX format.
            string filePath = "ResultDetectBorder.docx";
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}