using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Sample
{
    class Sample
    {
        
        static void Main(string[] args)
        {
            AddPictures();
        }

		/// <summary>
        /// How to add pictures into a document. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/add-pictures.php
        /// </remarks>
        public static void AddPictures()
        {
            string documentPath = @"Pictures.docx";
            string pictPath = @"..\..\..\image1.jpg";

            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();

            // Add a new section.
            Section s = new Section(dc);
            dc.Sections.Add(s);

            // 1. Picture with InlineLayout:

            // Create a new paragraph with picture.
            Paragraph par = new Paragraph(dc);
            s.Blocks.Add(par);
            par.ParagraphFormat.Alignment = HorizontalAlignment.Left;

            // Add some text content.
            par.Content.End.Insert("Shrek and Donkey ", new CharacterFormat() { FontName = "Calibri", Size = 16.0, FontColor = Color.Black });

            // Our picture has InlineLayout - it doesn't have positioning by coordinates
            // and located as flowing content together with text (Run and other Inline elements).
            Picture pict1 = new Picture(dc, InlineLayout.Inline(new Size(100, 100)), pictPath);

            // Add picture to the paragraph.
            par.Inlines.Add(pict1);

            // Add some text content.
            par.Content.End.Insert(" arrive at Farquaad's palace in Duloc, where they end up in a tournament.", new CharacterFormat() { FontName = "Calibri", Size = 16.0, FontColor = Color.Black });

            // 2. Picture with FloatingLayout:
            // Floating layout means that the Picture (or Shape) is positioned by coordinates.
            Picture pict2 = new Picture(dc, pictPath);
            pict2.Layout = FloatingLayout.Floating(
                new HorizontalPosition(50, LengthUnit.Millimeter, HorizontalPositionAnchor.Page),
                new VerticalPosition(70, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new Size(LengthUnitConverter.Convert(10, LengthUnit.Centimeter, LengthUnit.Point),
                         LengthUnitConverter.Convert(10, LengthUnit.Centimeter, LengthUnit.Point))
                         );

            // Set the wrapping style.
            (pict2.Layout as FloatingLayout).WrappingStyle = WrappingStyle.Square;

            // Add our picture into the section.
            s.Content.End.Insert(pict2.Content);

            // Save our document into DOCX format.
            dc.Save(documentPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}