using System;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            ShowInlines();
        }
        /// <summary>
        /// Iterates through a document and count the amount of Paragraphs and Runs.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/elementcollection-iteration.php
        /// </remarks>
        static void ShowInlines()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            Console.WriteLine("This document contains from:");
            for (int sect = 0; sect < dc.Sections.Count; sect++)
            {
                Console.WriteLine("Section {0} contains from:", sect);
                int totalParagraphs = 0;
                Section section = dc.Sections[sect];
                for (int blocks = 0; blocks < section.Blocks.Count; blocks++)
                {
                    if (section.Blocks[blocks] is Paragraph)
                    {
                        totalParagraphs++;
                        Paragraph paragraph = section.Blocks[blocks] as Paragraph;
                        Console.Write("\t\t Paragraph {0} contains from: ", totalParagraphs);
                        int totalRuns = 0;
                        for (int i = 0; i < paragraph.Inlines.Count; i++)
                        {
                            if (paragraph.Inlines[i] is Run)
                                totalRuns++;
                        }
                        Console.WriteLine("{0} Run(s).", totalRuns);
                    }


                }
            }
            Console.ReadKey();
        }
    }
}