using System;
using System.Threading.Tasks;
using SautinSoft.Document;
using SautinSoft.Document.CustomMarkups;
using SautinSoft.Document.Drawing;
using SautinSoft.Document.Tables;
using System.Net;
using System.IO;


namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertPicture();
        }
        /// <summary>
        /// Inserting a Picture content control.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/content-controls-insert-picture-net-csharp-vb.php
        /// </remarks>

        static void InsertPicture()
        {
            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();
            Picture pict, pict1, pict2;

            byte[] imageBytes = File.ReadAllBytes(@"..\..\..\banner_sautinsoft.jpg");
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                pict = new Picture(dc, new InlineLayout(new Size(400, 100)), ms);
            }

            imageBytes = File.ReadAllBytes(@"..\..\..\developer.png");
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                pict1 = new Picture(dc, new InlineLayout(new Size(50, 50)), ms);
            }

            Section section = new Section(dc);
            dc.Sections.Add(section);

            section.Blocks.Add(new Paragraph(dc, "Picture below is inside the block-level content control:"));

            // Create a picture content control.
            BlockContentControl control = new BlockContentControl(dc, ContentControlType.Picture, new Paragraph(dc, pict));
            section.Blocks.Add(control);

            Paragraph par = new Paragraph(dc,
                new Run(dc, "Following picture is inside the inline-level content control: "),
                new InlineContentControl(dc, ContentControlType.Picture, pict1));
            section.Blocks.Add(par);

            section.Blocks.Add(new Paragraph(dc, "Insert a picture content control from the local disk:"));
            string pictPath = @"..\..\..\picture.jpg";
            pict2 = new Picture(dc, new InlineLayout(new Size(100, 100)), pictPath);
            BlockContentControl localpict = new BlockContentControl(dc, ContentControlType.Picture, new Paragraph(dc, pict2));
            section.Blocks.Add(localpict);

            // Save our document into DOCX format.
            string resultPath = @"result.docx";
            dc.Save(resultPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}

