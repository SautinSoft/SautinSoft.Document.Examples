using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Sample
{
    class Sample
    {
       
        static void Main(string[] args)
        {
            Shapes();
        }
        
		/// <summary>
        /// This sample shows how to work with shapes. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/shapes.php
        /// </remarks>
        public static void Shapes()
        {
            string documentPath = @"Shapes.docx";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();

            // Create shape 1 with fill and outline.
            Shape shp1 = new Shape(dc, Layout.Floating(
                new HorizontalPosition(25f, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin),
                new VerticalPosition(20f, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new Size(200, 100)
                ));

            // Specify outline and fill using a picture.
            shp1.Outline.Fill.SetSolid(Color.DarkGreen);
            shp1.Outline.Width = 2;

            // Set fill for this shape.
            shp1.Fill.SetSolid(Color.Orange);

            // Create shape 2 with some text inside, 100mm*20mm.
            Shape shp2 = new Shape(dc, Layout.Floating(
                new HorizontalPosition(100f, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin),
                new VerticalPosition(20f, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new Size(LengthUnitConverter.Convert(100f, LengthUnit.Millimeter, LengthUnit.Point),
                        LengthUnitConverter.Convert(20f, LengthUnit.Millimeter, LengthUnit.Point))
                ));

            // Specify outline and fill using a picture.
            shp2.Outline.Fill.SetSolid(Color.LightGray);
            shp2.Outline.Width = 0.5;

            // Create a new paragraph with a formatted text.
            Paragraph p = new Paragraph(dc);
            Run run1 = new Run(dc, "Welcome to International Software Developer conference!");
            run1.CharacterFormat.FontName = "Helvetica";
            run1.CharacterFormat.Size = 14f;
            run1.CharacterFormat.Italic = true;
            p.Inlines.Add(run1);

            // Add the paragraph into the shp2.Text property.
            shp2.Text.Blocks.Add(p);

            // Add our shapes into the document.
            dc.Content.End.Insert(shp1.Content);
            dc.Content.End.Insert(shp2.Content);

            // Save the document to DOCX format.
            dc.Save(documentPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}