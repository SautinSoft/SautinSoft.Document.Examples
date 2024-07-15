using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            // Get your free 100 - day key here:   
            // https://sautinsoft.com/start-for-free/

            FullPageImage();
        }

        /// <summary>
        /// How to add pictures into a document. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/how-to-map-image-csharp-net.php
        /// </remarks>
        public static void FullPageImage()
        {
            string documentPath = @"pictures.docx";
            string pictPath = @"..\..\..\image-png.png";

            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();

            // Add a new section, A5 Landscape, and custom page margins.
            Section s = new Section(dc);
            s.PageSetup.PaperType = PaperType.A5;
            s.PageSetup.Orientation = Orientation.Landscape;
            s.PageSetup.PageMargins = new PageMargins()
            {
                Top = LengthUnitConverter.Convert(0, LengthUnit.Millimeter, LengthUnit.Point),
                Right = LengthUnitConverter.Convert(0, LengthUnit.Inch, LengthUnit.Point),
                Bottom = LengthUnitConverter.Convert(0, LengthUnit.Millimeter, LengthUnit.Point),
                Left = LengthUnitConverter.Convert(0, LengthUnit.Inch, LengthUnit.Point)
            };
            dc.Sections.Add(s);

            // Create a new paragraph with picture.
            Paragraph par = new Paragraph(dc);
            s.Blocks.Add(par);

            // Our picture has InlineLayout - it doesn't have positioning by coordinates
            // and located as flowing content together with text (Run and other Inline elements).
            Picture pict1 = new Picture(dc, InlineLayout.Inline(new Size(s.PageSetup.PageWidth, s.PageSetup.PageHeight)), pictPath);

            // Add picture to the paragraph.
            par.Inlines.Add(pict1);

            // Save our document into DOCX format.
            dc.Save(documentPath);

            // Save our document into PDF format.
            dc.Save(@"PdfDocument.pdf");

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(@"PdfDocument.pdf") { UseShellExecute = true });
        }
    }
}