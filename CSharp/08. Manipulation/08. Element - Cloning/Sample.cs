using System;
using SautinSoft.Document;
using System.IO;
using System.Linq;
using System.Text;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            CloningElement();
        }

        /// <summary>
        /// How to clone an element in DOCX document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/cloning-element-net-csharp-vb.php
        /// </remarks>
        static void CloningElement()
        {
            string filePath = @"..\..\..\Parsing.docx";
            string cloningFile = "Cloning.docx";
            DocumentCore dc = DocumentCore.Load(filePath);

            // Clone section.
            dc.Sections.Add(dc.Sections[0].Clone(true));

            // Clone paragraphs.
            foreach (Block item in dc.Sections[0].Blocks)
                dc.Sections.Last().Blocks.Add(item.Clone(true));

            // Save the result.
            dc.Save(cloningFile);

            // Show the result.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(cloningFile) { UseShellExecute = true });
        }
    }
}