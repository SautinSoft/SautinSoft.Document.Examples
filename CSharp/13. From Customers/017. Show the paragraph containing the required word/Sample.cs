using System;
using System.IO;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            {
                FindWordInParagraph();
            }
        }
        /// <summary>
        /// Find any "word" in a folder with PDF files inside and show a paragraph, where this word will be found.
        /// You may change the extension: pdf, docx, rtf.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-show-paragraph-containing-required-word-in-csharp-vb-net.php
        /// </remarks>
        static void FindWordInParagraph()
        {
            // A regular expression (shortened as regex or regexp; sometimes referred to as rational expression) is a sequence of characters that specifies a search pattern in text.
            Regex regex = new Regex(@"\bcompany\b", RegexOptions.IgnoreCase);

          
            string filePath = @"..\..\..\instance.pdf";
            
                DocumentCore dc = DocumentCore.Load(filePath);

                // Provides a functionality to paginate the document content.
                DocumentPaginator dp = dc.GetPaginator();
                foreach (ContentRange content in dc.Content.Find(regex))
                {
                    ElementFrame ef = dp.GetElementFrames().FirstOrDefault(e => content.Start.Equals(e.Content.Start));
                    Paragraph paragraph = content.Start.Parent.Parent as Paragraph;

                    // We are looking for a sentence in which this word was found.
                    string sentence = paragraph.Content.ToString().Trim();
                    Console.WriteLine("Filename: " + filePath + "\r\n" + sentence);

                    // The coordinates of the found word.
                    Console.WriteLine("Info:" + ef.Bounds.ToString());
                    Console.WriteLine("Next paragraph?");
                    Console.ReadKey();
                }   
        }
    }
}
