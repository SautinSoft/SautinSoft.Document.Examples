using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Sample
{
    class Sample
    {
       
        static void Main(string[] args)
        {
            HeadersAndFooters();
        }

		/// <summary>
        /// How to add a header and footer into a document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/headersfooters.php
        /// </remarks>
        public static void HeadersAndFooters()
        {
            string documentPath = @"HeadersAndFooters.docx";

            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();

            // Add a new section in the document.
            Section s = new Section(dc);
            dc.Sections.Add(s);

            // Let's add a paragraph with text.
            Paragraph p = new Paragraph(dc);
            dc.Sections[0].Blocks.Add(p);

            p.ParagraphFormat.Alignment = HorizontalAlignment.Justify;
            p.Content.Start.Insert("Once upon a time, in a far away swamp, there lived an ogre named Shrek whose precious " +
                                   "solitude is suddenly shattered by an invasion of annoying fairy tale characters...", new CharacterFormat() { Size = 12, FontName = "Arial" });

            // Create a new header with formatted text.
            HeaderFooter header = new HeaderFooter(dc, HeaderFooterType.HeaderDefault);
            header.Content.Start.Insert("Shrek and Donkey travel to the castle and split up to find Fiona.", new CharacterFormat() { Size = 14.0 });

            // Add the header into HeadersFooters collection of the 1st section.
            s.HeadersFooters.Add(header);

            // Create a new footer with some text and image.
            HeaderFooter footer = new HeaderFooter(dc, HeaderFooterType.FooterDefault);

            // Create a paragraph to insert it into the footer.
            Paragraph par = new Paragraph(dc);
            par.Content.Start.Insert("Shrek and Donkey travel to the castle and split up to find Fiona. ", new CharacterFormat() { Size = 14.0 });
            par.ParagraphFormat.Alignment = HorizontalAlignment.Left;

            // Insert image into the paragraph.
            double wPt = LengthUnitConverter.Convert(7, LengthUnit.Centimeter, LengthUnit.Point);
            double hPt = LengthUnitConverter.Convert(7, LengthUnit.Centimeter, LengthUnit.Point);

            Picture pict = new Picture(dc, Layout.Inline(new Size(wPt, hPt)), @"..\..\..\image1.jpg");
            par.Inlines.Add(pict);

            // Add the paragraph into the Blocks collection of the footer.
            footer.Blocks.Add(par);

            // Finally, add the footer into 1st section (HeadersFooters collection).
            s.HeadersFooters.Add(footer);

            // Save the document into DOCX format.
            dc.Save(documentPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}