using System;
using SautinSoft.Document;
using System.Text;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertingImage();
        }
        /// <summary>
        /// Insert an image and shape inline or in the specified position using DocumentBuilder.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-image.php
        /// </remarks>

        static void InsertingImage()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            string resultPath = @"Result.docx";
            string pictPath = @"..\..\..\logo.png";

            // Insert the formatted text into the document using DocumentBuilder.
            db.CharacterFormat.FontName = "Courier";
            db.CharacterFormat.Size = 17f;
            db.CharacterFormat.Italic = true;
            db.CharacterFormat.FontColor = Color.Orange;
            db.Writeln("Insert an Image and Shape using DocumentBuilder.");

            // Images:
            // 1st way: Insert an Inline image into the document.
            // Specify the image size and rotation (if required).
            Picture pict1 = db.InsertImage(pictPath, new Size(100, 30, LengthUnit.Millimeter));
            pict1.Rotation = -3;

            // 2nd way: Insert a Floating image from a file at the specified position.
            Picture pict2 = db.InsertImage(pictPath, new HorizontalPosition(1, LengthUnit.Centimeter, HorizontalPositionAnchor.LeftMargin),
                 new VerticalPosition(8, LengthUnit.Centimeter, VerticalPositionAnchor.TopMargin), WrappingStyle.InFrontOfText);

            // Shapes:
            // 1st way: Insert an Inline shape.
            Shape shp1 = db.InsertShape(Figure.SmileyFace, new Size(3, 3, LengthUnit.Centimeter));

            // 2nd way: Insert a Floating shape with specified position, size and text wrap style.
            Size size1 = new Size(7, 6, LengthUnit.Centimeter);
            Shape shp2 = db.InsertShape(Figure.RoundRectangle, new HorizontalPosition(8, LengthUnit.Centimeter,HorizontalPositionAnchor.LeftMargin), 
                new VerticalPosition(10, LengthUnit.Centimeter, VerticalPositionAnchor.TopMargin), WrappingStyle.InFrontOfText, new Size(7, 6, LengthUnit.Centimeter));

            shp2.Fill.SetSolid(Color.White);

            // Move the "cursor" position inside the shape content.
            db.MoveTo(shp2.Text.Blocks.Content.Start);
            db.CharacterFormat.FontColor = Color.Green;
            db.CharacterFormat.Size = 26f;
            db.Writeln("Text inside Shape.");

            // Move the "cursor" back to the paragraph with "shp1".
            db.MoveTo(shp1.Content.End);            

            // Save our document into DOCX format.
            dc.Save(resultPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}