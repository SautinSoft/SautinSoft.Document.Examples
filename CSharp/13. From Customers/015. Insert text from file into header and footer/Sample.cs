using System;
using System.IO;
using System.Collections.Generic;
using SautinSoft.Document;
using System.Linq;
using SautinSoft.Document.Tables;
using System.Text;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free 100-day key here:   
            // https://sautinsoft.com/start-for-free/

            InsertTextHeaderFooter();
        }

        /// <summary>
        /// How to insert the contents of a file into the header and footer of an existing HTML.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-insert-text-from-file-into-header-footer-in-csharp-vb-net.php
        /// </remarks>

        static void InsertTextHeaderFooter()
        {   
            string inpFile = @"..\..\..\example.html";
            string outFile = @"result.docx";
            
            // Reads all text from HTML file.
            byte[] htmlHeaderBytes = Encoding.UTF8.GetBytes(File.ReadAllText(@"..\..\..\header.html"));
            string htmlHeader = System.Text.Encoding.UTF8.GetString(htmlHeaderBytes);

            // Reads all text from RTF file.
            byte[] rtfFooterBytes = Encoding.UTF8.GetBytes(File.ReadAllText(@"..\..\..\footer.rtf"));
            string rtfFooter = System.Text.Encoding.UTF8.GetString(rtfFooterBytes);

            // Load a document.
            DocumentCore dc = DocumentCore.Load(inpFile);
            
            // Create a new header with formatted HTML text.
            HeaderFooter header = new HeaderFooter(dc, HeaderFooterType.HeaderDefault);
            
            // Add the header into HeadersFooters collection and Clone to all sections.
            header.Content.Start.Insert(htmlHeader, LoadOptions.HtmlDefault);
            foreach (Section s in dc.Sections)
                {
                    s.HeadersFooters.Add(header.Clone(true));
                }

            // Create a new footer with formatted RTF text.
            HeaderFooter footer = new HeaderFooter(dc, HeaderFooterType.FooterDefault);

            // Add the footer into HeadersFooters collection and Clone to all sections.
            footer.Content.Start.Insert(rtfFooter, LoadOptions.RtfDefault);
            foreach (Section s in dc.Sections)
            {
                s.HeadersFooters.Add(footer.Clone(true));
            }

            // Save the result as DOCX file.
            dc.Save(outFile);

                // Open the result for demonstration purposes.
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
            }
        }
    }