using System;
using System.IO;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            MailMergeSimpleEnvelope();
        }

        /// <summary>
        /// Generates 5 envelopes "Happy New Year" for Simpson family using the one template.
        /// </summary>
        /// <remarks>
        /// See details at: https://sautinsoft.com/products/document/help/net/developer-guide/mail-merge-simple-report-net-csharp-vb.php
        /// </remarks>
        public static void MailMergeSimpleEnvelope()
        {
            string templatePath = @"..\..\..\envelope-template.docx";
            string resultPath = "Simpson-family.docx";

            DocumentCore dc = DocumentCore.Load(templatePath);

            var dataSource = new[] { new { Name = "Homer", FamilyName = "Simpson" },
                                new { Name = "Marge ", FamilyName = "Simpson" },
                                new { Name = "Bart", FamilyName = "Simpson" },
                                new { Name = "Lisa", FamilyName = "Simpson" },
                                new { Name = "Maggie", FamilyName = "Simpson" }};

            dc.MailMerge.Execute(dataSource);
            dc.Save(resultPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}
