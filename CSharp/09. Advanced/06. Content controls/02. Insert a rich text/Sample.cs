using System;
using System.Text;
using SautinSoft.Document;
using SautinSoft.Document.CustomMarkups;
using SautinSoft.Document.Drawing;
using System.IO;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertRichText();
        }
        /// <summary>
        /// Inserting a Rich text content control.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/content-controls-insert-rich-text-net-csharp-vb.php
        /// </remarks>

        static void InsertRichText()
        {
            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();

            // Create a rich text content control.
            BlockContentControl rt = new BlockContentControl(dc, ContentControlType.RichText);
            dc.Sections.Add(new Section(dc, rt));

            // Add the content control properties.
            rt.Properties.Title = "Rich text content control.";
            rt.Properties.Color = Color.Blue;
            rt.Properties.LockDeleting = true;
            rt.Properties.LockEditing = true;
            rt.Document.DefaultCharacterFormat.FontColor = Color.Orange;

            // Add new paragraphs with formatted text.
            rt.Blocks.Add(new Paragraph(dc, "Line 1"));
            rt.Blocks.Add(new Paragraph(dc, "Line 2"));
            rt.Blocks.Add(new Paragraph(dc, "Line 3"));

            // Add a picture and shape to the block.
            string pictPath = @"..\..\..\banner_sautinsoft.jpg";
            Picture pict = new Picture(dc, InlineLayout.Inline(new Size(400, 100)), pictPath);
            Shape shp = new Shape(dc, Layout.Inline(new Size(4, 1, LengthUnit.Centimeter)));
            Paragraph im = new Paragraph(dc);
            rt.Blocks.Add(im);
            im.Inlines.Add(pict);
            Paragraph sh = new Paragraph(dc);
            rt.Blocks.Add(sh);
            sh.Inlines.Add(shp);

            rt.Blocks.Add(new Paragraph(dc, "Line 4"));
            
            // Save our document into DOCX format.
            string resultPath = @"result.docx";
            dc.Save(resultPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}
