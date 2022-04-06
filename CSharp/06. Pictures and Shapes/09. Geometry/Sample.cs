using System;
using System.IO;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Sample
{
    class Sample
    {
       
        static void Main(string[] args)
        {
            Geometry();
        }
     
		/// <summary>
        /// This sample shows how to work with shapes and geometry. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/geometry.php
        /// </remarks>
        public static void Geometry()
        {
            string pictPath = @"..\..\..\image1.jpg";
            string documentPath = @"Geometry.docx";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();

            // Create shape 1 with preset geometry (Smiley Face).
            Shape shp1 = new Shape(dc, Layout.Floating(
                new HorizontalPosition(20f, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin),
                new VerticalPosition(80f, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new Size(100, 100)
                ));

            // Specify outline and fill.
            shp1.Outline.Fill.SetSolid(new Color("358CCB"));
            shp1.Outline.Width = 3;
            shp1.Fill.SetSolid(Color.Orange);

            // Specify a figure.
            shp1.Geometry.SetPreset(Figure.SmileyFace);

            // Create shape 2 with custom geometry path (using points array).
            Shape shp2 = new Shape(dc, Layout.Floating(
                new HorizontalPosition(85f, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin),
                new VerticalPosition(80f, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new Size(100, 100)
                ));

            // Specify outline and fill using a picture.
            shp2.Outline.Fill.SetSolid(Color.Green);
            shp2.Outline.Width = 2;

            // Set the picture as fill for this shape.
            shp2.Fill.SetPicture(pictPath);

            // Specify the maximum X and Y coordinates that should be used
            // for within the path coordinate system.
            Size size = new Size(1, 1);

            // Specify the path points (draw a circle of 10 points).
            Point[] points = new Point[10];
            double a = 0;
            for (int i = 0; i < 10; ++i)
            {
                points[i] = new Point(0.5 + Math.Sin(a) * 0.5, 0.5 + Math.Cos(a) * 0.5);
                a += 2 * Math.PI / 10;
            }

            // Create and add new custom path from specified points array.
            shp2.Geometry.SetCustom().AddPath(size, points, true);

            // Create shape3 with custom geometry path (using path elements).
            Shape shp3 = new Shape(dc, Layout.Floating(
                new HorizontalPosition(150f, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin),
                new VerticalPosition(80f, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new Size(100, 100)
                ));

            // Specify outline and fill.
            shp3.Outline.Fill.SetSolid(new Color(255,0,0));
            shp3.Outline.Width = 2;
            shp3.Fill.SetSolid(Color.Yellow);

            // Create and add new custom path.
            CustomPath path = shp3.Geometry.SetCustom().AddPath(new Size(1, 1));

            // Specify path elements.
            path.MoveTo(new Point(0, 0));
            path.AddLine(new Point(0, 1));
            path.AddLine(new Point(1, 1));
            path.AddLine(new Point(1, 0));
            path.ClosePath();

            // Add drawing elements to document.
            dc.Content.End.Insert(shp1.Content);
            dc.Content.End.Insert(shp2.Content);
            dc.Content.End.Insert(shp3.Content);

            // Save the document to DOCX format.
            dc.Save(documentPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}