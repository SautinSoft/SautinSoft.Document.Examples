using System.Linq;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        
        static void Main(string[] args)
        {
            InsertContent();
        }

		/// <summary>
        /// Create a new document with some content.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-content-net-csharp-vb.php
        /// </remarks>
        public static void InsertContent()
        {
            string documentPath = @"InsertContent.pdf";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();

            // Add new section.
            Section section = new Section(dc);
            dc.Sections.Add(section);

            // You may add a new paragraph using a classic way:
            section.Blocks.Add(new Paragraph(dc, "This is the first line in 1st paragraph!"));
            Paragraph par1 = section.Blocks[0] as Paragraph;
            par1.ParagraphFormat.Alignment = HorizontalAlignment.Center;

            // But here, let's see how to do it using ContentRange:

            // 1. Insert a content as Text.
            dc.Content.End.Insert("\nThis is a first line in 2nd paragraph.", new CharacterFormat() { Size = 25, FontColor = Color.Blue, Bold = true });

            // 2. Insert a content as HTML (at the start).
            dc.Content.Start.Insert("Hello from HTML!", new HtmlLoadOptions());

            // 3. Insert a content as RTF.
            dc.Content.End.Insert(@"{\rtf1 \b The line from RTF\b0!\par}", new RtfLoadOptions());

            // 4. Insert a content of SpecialCharacter.
            dc.Content.End.Insert(new SpecialCharacter(dc, SpecialCharacterType.LineBreak).Content);

            // 5. Insert a content as ContentRange.
            // Find first content of "HTML" and insert it at the end.
            ContentRange cr = dc.Content.Find("HTML").First();
            dc.Content.End.Insert(cr);

            // 6. Insert the content of Paragraph at the end.
            Paragraph p = new Paragraph(dc);
            Run run = new Run(dc, "As summarize, you can insert content of any class " +
                "derived from Element: Table, Shape, Paragraph, Run, HeaderFooter and even a whole DocumentCore!");
            run.CharacterFormat.FontColor = Color.DarkGreen;
            p.Inlines.Add(run);
            dc.Content.End.Insert(p.Content);

            // Save our document into PDF format.
            dc.Save(documentPath, new PdfSaveOptions() { Compliance = PdfCompliance.PDF_A1a });

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}