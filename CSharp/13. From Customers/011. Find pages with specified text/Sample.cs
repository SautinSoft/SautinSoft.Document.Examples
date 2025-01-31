using SautinSoft.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            FindPagesSpecifiedText();       
        }
        /// <summary>
        /// How to find out on which pages of the document the required word is located.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-find-pages-with-specified-text-net-csharp-vb.php
        /// </remarks>
        public static void FindPagesSpecifiedText()
        {
            // The path for input files or directory.
            string inpFile = @"..\..\..\example.docx";

            // What we need to search.
            var searchText = "Invoice";
            int quantity = 0;
            
            // Load our documument in Document's engine.
            DocumentCore dc = DocumentCore.Load(inpFile);
            
            // Regex https://en.wikipedia.org/wiki/Regular_expression
            Regex regex = new Regex(searchText, RegexOptions.IgnoreCase);

            // Document paginator allows you to calculate of pages.
            DocumentPaginator dp = dc.GetPaginator();
            
            // We will search "searchText" on each pages (enumeration).
            for (int page = 0; page < dp.Pages.Count; page++)
            {
                foreach (ContentRange item in dp.Pages[page].Content.Find(regex).Reverse())
                {
                    Console.WriteLine($"I see the [{searchText}] on the page # {page + 1}");
                    quantity++;
                }
            }
            Console.WriteLine();
            Console.WriteLine($"I met [{searchText}] {quantity} times.  Please click on any button");
            Console.ReadKey();
        }
    }
}