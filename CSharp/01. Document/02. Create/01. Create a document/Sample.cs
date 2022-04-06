using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // You can create the same document by using 3 ways:
            //
            //  + DocumentBuilder
            //  + DOM directly
            //  + DOM and ContentRange
            //
            // Choose any of them which you like.

            // Way 1:
            CreateUsingDocumentBuilder();

            // Way 2:
            CreateUsingDOM();

            // Way 3:
            CreateUsingContentRange();
        }

        /// <summary>
        /// Creates a new document using DocumentBuilder and saves it in a desired format.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/create-document.php
        /// </remarks>
        static void CreateUsingDocumentBuilder()
        {
            // Create a new document and DocumentBuilder.
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            // Specify the formatting and insert text.
            db.CharacterFormat.FontName = "Verdana";
            db.CharacterFormat.Size = 65.5f;
            db.CharacterFormat.FontColor = Color.Orange;
            db.Write("Hello World!");

            // Save the document in DOCX format.
            string outFile = "DocumentBuilder.docx";
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
        /// <summary>
        /// Creates a new document using DOM and saves it in a desired format.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/create-document.php
        /// </remarks>
        static void CreateUsingDOM()
        {
            // Create a new document.
            DocumentCore dc = new DocumentCore();

            // Create a new section,
            // add the section the document.
            Section sect = new Section(dc);
            dc.Sections.Add(sect);

            // Create a new paragraph,
            // add the paragraph to the section.
            Paragraph par = new Paragraph(dc);
            sect.Blocks.Add(par);

            // Create a new run (text object),
            // add the run to the paragraph.
            Run run = new Run(dc, "Hello World!");
            run.CharacterFormat.FontName = "Verdana";
            run.CharacterFormat.Size = 65.5f;
            run.CharacterFormat.FontColor = Color.Orange;
            par.Inlines.Add(run);

            // Save the document in PDF format.
            string outFile = @"DOM.pdf";
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }


        /// <summary>
        /// Creates a new document using DOM and ContentRange and saves it in a desired format.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/create-document.php
        /// </remarks>
        static void CreateUsingContentRange()
        {
            // Create a new document.
            DocumentCore dc = new DocumentCore();
            // Insert the formatted text into the document.
            dc.Content.End.Insert("Hello World!", new CharacterFormat() { FontName = "Verdana", Size = 65.5f, FontColor = Color.Orange });

            // Save the document in HTML format.
            string outFile = @"ContentRange.html";
            dc.Save(outFile, new HtmlFixedSaveOptions() { Title = "ContentRange" });

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}