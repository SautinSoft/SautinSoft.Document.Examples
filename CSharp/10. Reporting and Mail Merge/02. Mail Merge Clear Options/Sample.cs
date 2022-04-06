using SautinSoft.Document;
using SautinSoft.Document.MailMerging;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            MailMergeWithClearOptions();
        }
        /// <summary>
        /// Shows how use ClearOptions - remove specific elements if no data has been imported into them.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/mail-merge-clear-options-csharp-vb.php
        /// </remarks>
        static void MailMergeWithClearOptions()
        {
            DocumentCore document = DocumentCore.Load(@"..\..\..\MailMergeClearOptions.docx");

            // Example 1: Remove fields for which no data has been found in the mail merge data source.
            var dataSource1 = new
            {
                Field2 = "Some text"
            };
            document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveUnusedFields;
            document.MailMerge.Execute(dataSource1, "Example1");


            // Example 2: Remove paragraphs contained merge fields but none of them has been merged.
            var dataSource2 = new
            {
                Field1 = string.Empty,
                Field2 = "Some text"
            };

            document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyParagraphs;
            document.MailMerge.Execute(dataSource2, "Example2");


            // Example 3: Remove table rows contained merge fields but none of them has been merged.
            var dataSource3 = new
            {
                Field1 = "Some text 1",
                Field2 = (string)null,
                Field3 = "Some text 3",
            };

            document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyTableRows;
            document.MailMerge.Execute(dataSource3, "Example3");


            // Example 4: Remove ranges into which no field has been merged.
            var dataSource4 = new
            {
                TotalRecords = 2,
                Records = new object[]
                {
                    new
                    {
                        Record = 1,
                        Text = "Some text 1"
                    },
                    new
                    {
                        Record = (object)null,
                        Text = string.Empty
                    },
                    new
                    {
                        Record = 3,
                        Text = "Some text 3"
                    },
            }
            };
            document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyRanges;
            document.MailMerge.Execute(dataSource4, "Example4");

            string resultPath = "ClearOptions.docx";

            // Save the output to file.
            document.Save(resultPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}