using System;
using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            DifferentHeaderAndFooters();
        }
        /// <summary>
        /// Creates a document with different headers: on first page, default and in another section.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-docx-document-with-different-headers-net-csharp-vb.php
        /// </remarks>
        static void DifferentHeaderAndFooters()
        {
            string documentPath = "DifferentHeaders.docx";

            DocumentCore dc = new DocumentCore();
            Section section1 = new Section(dc);
            dc.Sections.Add(section1);

            // Create headers.
            HeaderFooter defaultHeader = CreateDefaultHeader(dc);
            HeaderFooter firstHeader = CreateFirstHeader(dc);

            // Add the headers into the collection.
            section1.HeadersFooters.Add(defaultHeader);
            section1.HeadersFooters.Add(firstHeader);

            // Set that main section has a different header on the first page.
            section1.PageSetup.TitlePage = true;

            // Add some content.
            Paragraph p = new Paragraph(dc, "My father\'s family name being Pirrip, and my Christian name Philip, " +
                "my infant tongue could make of both names nothing longer or more explicit than Pip. So, I called myself " +
                "Pip, and came to be called Pip. I give Pirrip as my father\'s family name, on the authority of his " +
                "tombstone and my sister, -Mrs. Joe Gargery, who married the blacksmith.");
            int totalPars = 15;

            for (int i = 0; i < totalPars; i++)
            {
                section1.Blocks.Add(p.Clone(true));
            }

            // Add another header and content in the document.

            // Let's create a one more section.
            Section section2 = new Section(dc);
            dc.Sections.Add(section2);

            // Create a new header.
            HeaderFooter header = new HeaderFooter(dc, HeaderFooterType.HeaderDefault);
            header.Content.Start.Insert("Chapter II - another header.", new CharacterFormat()
            {
                FontColor = Color.Blue,
                Size = 18
            });

            // Add the header into the Section2.
            section2.HeadersFooters.Add(header);

            // Add some content.
            section2.Content.Start.Insert("My sister, Mrs. Joe Gargery, was more than twenty years older than I, " +
                "and had established a great reputation with herself and the neighbors because she had brought me " +
                "up \"by hand.\" Having at that time to find out for myself what the expression meant, and knowing " +
                "her to have a hard and heavy hand, and to be much in the habit of laying it upon her husband as well " +
                "as upon me, I supposed that Joe Gargery and I were both brought up by hand.");

            // Save the document into DOCX format.
            dc.Save(documentPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
        static HeaderFooter CreateDefaultHeader(DocumentCore dc)
        {
            HeaderFooter header = new HeaderFooter(dc, HeaderFooterType.HeaderDefault);
            header.Content.Start.Insert("Chapter I - This is the default header", new CharacterFormat()
            {
                FontName = "Verdana",
                Size = 18.0,
                FontColor = Color.Orange
            });
            return header;
        }

        static HeaderFooter CreateFirstHeader(DocumentCore dc)
        {
            HeaderFooter header = new HeaderFooter(dc, HeaderFooterType.HeaderFirst);
            header.Content.Start.Insert("Charles Dickens. Great Expectations - 1st page header", new CharacterFormat()
            {
                FontName = "Verdana",
                Size = 18.0,
                FontColor = Color.DarkGreen
            });
            return header;
        }
    }
}