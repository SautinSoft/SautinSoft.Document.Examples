using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            PictureAndShape();
        }
        /// <summary>
        /// Creates a new document with shape containing a text and picture.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/pictures-and-shapes.php
        /// </remarks>
        static void PictureAndShape()
        {
            string filePath = @"Shape.docx";
            string imagePath = @"..\..\..\image.jpg";

            DocumentCore dc = new DocumentCore();                        
            
            // 1. Shape with text.
            Shape shapeWithText = new Shape(dc, Layout.Floating(new HorizontalPosition(1, LengthUnit.Inch, HorizontalPositionAnchor.Page), 
                new VerticalPosition(2, LengthUnit.Inch, VerticalPositionAnchor.Page),
                new Size(LengthUnitConverter.Convert(6, LengthUnit.Inch, LengthUnit.Point), LengthUnitConverter.Convert(1.5d, LengthUnit.Centimeter, LengthUnit.Point))));
            (shapeWithText.Layout as FloatingLayout).WrappingStyle = WrappingStyle.InFrontOfText;
            shapeWithText.Text.Blocks.Add(new Paragraph(dc, new Run(dc, "This is the text in shape.", new CharacterFormat() { Size  = 30})));
            shapeWithText.Outline.Fill.SetEmpty();
            shapeWithText.Fill.SetSolid(Color.Orange);
            dc.Content.End.Insert(shapeWithText.Content);            

            // 2. Picture with FloatingLayout:
            // Floating layout means that the Picture (or Shape) is positioned by coordinates.
            Picture pic = new Picture(dc, imagePath);
            pic.Layout = FloatingLayout.Floating(
                new HorizontalPosition(50, LengthUnit.Millimeter, HorizontalPositionAnchor.Page),
                new VerticalPosition(20, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new Size(LengthUnitConverter.Convert(10, LengthUnit.Centimeter, LengthUnit.Point),
                         LengthUnitConverter.Convert(10, LengthUnit.Centimeter, LengthUnit.Point))
                         );

            // Set the wrapping style.
            (pic.Layout as FloatingLayout).WrappingStyle = WrappingStyle.BehindText;

            // Add our picture into the section.
            dc.Content.End.Insert(pic.Content);

            dc.Save(filePath);

            // Show the result.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}