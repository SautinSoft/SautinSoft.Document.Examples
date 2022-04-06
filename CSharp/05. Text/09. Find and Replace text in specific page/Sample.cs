using System;
using System.IO;
using SautinSoft.Document;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            FindAndReplaceSpecificTextWay1();
            //FindAndReplaceSpecificTextWay2();
        }
        /// <summary>
        /// Find and replace a certain text only on the second page of the DOCX document.
        /// Way #1:
        /// The DOCX document is loaded, then the analysis of the number of pages inside (GetPaginator).
        /// DOCX is not a "paged" format and has to be paginated.
        /// There may be problems with the text in the tables, since the transfer to a new page is possible.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-text-in-specific-page-after-specific-text-net-csharp-vb.php
        /// </remarks>
        public static void FindAndReplaceSpecificTextWay1()
        {
            string inpFileWord = @"..\..\..\example.docx";
            string outFileWord = @"result1.docx";
            var searchText = "italic";

            DocumentCore dc = DocumentCore.Load(inpFileWord);
            Regex regex = new Regex(searchText, RegexOptions.IgnoreCase);

            DocumentPaginator dp = dc.GetPaginator();

            if (dp.Pages.Count > 2)
            {
                // Find and replace a certain text only on the second page of the DOCX document.
                // If you need the first page - 0, the third page - 2, etc.
                foreach (ContentRange item in dp.Pages[1].Content.Find(regex).Reverse())
                {
                    item.Replace("This word has been corrected", new CharacterFormat() { BackgroundColor = Color.Yellow, FontName = "Arial", Size = 16.0 });
                }
            }
            dc.Save(outFileWord, new DocxSaveOptions());
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(inpFileWord) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFileWord) { UseShellExecute = true });
        }

        /// <summary>
        /// Find and replace a certain text only on the second page of the DOCX document.
        /// Way #2:
        /// The DOCX document is loaded. DOCX is not a "paged" format and has to be paginated.
        /// We will convert DOCX to PDF and then to save in DOCX back.
        /// There may be problems with Formatting, because we are doing reverse format conversion.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-text-in-specific-page-after-specific-text-net-csharp-vb.php
        /// </remarks>
        public static void FindAndReplaceSpecificTextWay2()
        {
            string inpFileWord = @"..\..\..\example.docx";
            string tempPDFFile = @"example_temp.pdf";
            string outFileWord = @"result2.docx";
            var searchText = "italic";

            DocumentCore dc1 = DocumentCore.Load(inpFileWord);
            dc1.Save(tempPDFFile);
            DocumentCore dc2 = DocumentCore.Load(tempPDFFile);
            Regex regex = new Regex(searchText, RegexOptions.IgnoreCase);

            {
                for (int index = 0; index < dc2.Sections.Count; index++)
                {
                    var page = dc2.Sections[index];
                    if (dc2.Sections.Count > 2)
                    {
                        // Find and replace a certain text only on the second page of the DOCX document.
                        // If you need the first page - 0, the third page - 2, etc.
                        foreach (ContentRange item in dc2.Sections[1].Content.Find(regex).Reverse())
                        {
                            item.Replace("This word has been corrected", new CharacterFormat() { BackgroundColor = Color.Yellow, FontName = "Arial", Size = 16.0 });
                        }
                    }
                }
                dc2.Save(outFileWord);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(inpFileWord) { UseShellExecute = true });
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFileWord) { UseShellExecute = true });
            }
        }
    }
}