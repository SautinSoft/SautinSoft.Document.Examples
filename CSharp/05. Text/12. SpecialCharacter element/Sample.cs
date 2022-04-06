using System.Linq;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            DeletePageBreak();
        }
        /// <summary>
        /// Working with special characters in a document. How delete all page breaks in DOCX.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/special-character-text-in-docx-document-net-csharp-vb.php
        /// </remarks>
        static void DeletePageBreak()
        {
            string filePath = @"..\..\..\example.docx";
            string fileResult = @"Result.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            foreach (SpecialCharacter sc in dc.GetChildElements(true, ElementType.SpecialCharacter).Reverse())
            {
                if (sc.CharacterType == SpecialCharacterType.PageBreak)
                    sc.Parent.Content.Delete();
            }
            dc.Save(fileResult);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fileResult) { UseShellExecute = true });
        }
    }
}