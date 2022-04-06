using System;
using System.Globalization;
using System.IO;
using System.Linq;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            FindTextFromTable();
        }

        /// <summary>
        /// How to remove the rows with the specified text from a table.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-find-text-from-table-net-csharp-vb.php
        /// </remarks>
        public static void FindTextFromTable()
        {
            int longLiverMinYears = 90;

            string inpFile = @"..\..\..\example.docx";
            string outFile = Path.ChangeExtension(inpFile, ".pdf");

            // Load a document with a table containing various persons with different age.
            DocumentCore dc = DocumentCore.Load(inpFile);

            // Find a first table in the document.
            Table table = (Table)dc.GetChildElements(true, ElementType.Table).First();

            // Loop by the all rows from the end.
            // Find long-livers.
            bool isLongLiver = false;

            for (int r = table.Rows.Count - 1; r > 0; r--)
            {
                isLongLiver = false;

                // Take the 3rd cell with the birth date.
                TableCell tc = table.Rows[r].Cells[2];

                // Get the birth date.
                DateTime birthDate = DateTime.Now;
                if (DateTime.TryParse(tc.Content.ToString(), CultureInfo.CreateSpecificCulture("en-US"), DateTimeStyles.None, out birthDate))
                {
                    // Get the person age.
                    // Remove the row if the person isn't long-liver.
                    if (CalculateAge(birthDate) >= longLiverMinYears)
                        isLongLiver = true;
                }
                // Remove the row if it doesn't contain a long-liver.
                if (!isLongLiver)
                    table.Rows.RemoveAt(r);
            }

            // Save the document as PDF.
            dc.Save(outFile, new PdfSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
        private static int CalculateAge(DateTime dateOfBirth)
        {
            int age = 0;
            age = DateTime.Now.Year - dateOfBirth.Year;
            if (DateTime.Now.DayOfYear < dateOfBirth.DayOfYear)
                age = age - 1;
            return age;
        }
    }
}
