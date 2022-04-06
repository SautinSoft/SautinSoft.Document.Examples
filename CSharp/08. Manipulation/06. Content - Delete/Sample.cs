using System.IO;
using System.Linq;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        
        static void Main(string[] args)
        {
            DeleteContent();
        }

		/// <summary>
        /// Open a document and delete some content.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/delete-content-net-csharp-vb.php
        /// </remarks>
        public static void DeleteContent()
        {
            string loadPath = @"..\..\..\example.docx";
            string savePath = "Result.docx";

            DocumentCore dc = DocumentCore.Load(loadPath);

            // Remove the text "This" from all paragraphs in 1st section.
            foreach (Paragraph par in dc.Sections[0].GetChildElements(true, ElementType.Paragraph))
            {
                var findText = par.Content.Find("This");

                if (findText != null)
                {
                    foreach (ContentRange cr in findText)
                        cr.Delete();
                }
            }
            dc.Save(savePath);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(loadPath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(savePath) { UseShellExecute = true });
        }
    }
}