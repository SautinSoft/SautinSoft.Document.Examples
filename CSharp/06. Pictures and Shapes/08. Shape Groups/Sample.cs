using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Sample
{
    class Sample
    {
       
        static void Main(string[] args)
        {
            ShapeGroups();
        }
     
		/// <summary>
        /// This sample shows how to work with shape groups. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/shape-groups.php
        /// </remarks>
        public static void ShapeGroups()
        {
            string pictPath = @"..\..\..\image1.jpg";
            string documentPath = @"ShapeGroups.docx";

            // Let's create a document.
            DocumentCore dc = new DocumentCore();

            // Create floating layout.
            HorizontalPosition hp = new HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Page);
            VerticalPosition vp = new VerticalPosition(5f, LengthUnit.Centimeter, VerticalPositionAnchor.TopMargin);
            FloatingLayout fl = new FloatingLayout(hp, vp, new Size(300, 300));

            // Create group.
            ShapeGroup group = new ShapeGroup(dc, fl);

            // Specify the size dimensions of the child extents rectangle.
            group.ChildSize = new Size(100, 100);

            // Create a child shape#1 (inside group) with preset geometry.
            // Specify shape's size and offset relative to group's ChildSize (100x100).
            Shape shape1 = new Shape(dc, new GroupLayout(new Point(0, 0), new Size(50, 50)));

            // Specify outline and fill.
            shape1.Outline.Fill.SetSolid(new Color("#358CCB"));
            shape1.Outline.Width = 2;
            shape1.Fill.SetSolid(Color.Orange);

            // Shape will be rectangle.
            shape1.Geometry.SetPreset(Figure.Rectangle);

            // Create picture and add it into the group.
            Picture picture = new Picture(dc, Layout.Group(new Point(50, 50), new Size(50, 50)), pictPath);

            // Specify picture fill mode.
            picture.ImageData.FillMode = PictureFillMode.Stretch;

            // Add shape and picture into our group.
            group.ChildShapes.Add(shape1);
            group.ChildShapes.Add(picture);

            // Add our group into the document.
            dc.Content.End.Insert(group.Content);

            // Save our document into DOCX format.
            dc.Save(documentPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}