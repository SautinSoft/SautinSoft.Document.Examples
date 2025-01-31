using System;
using System.IO;
using SautinSoft.Document;
using SkiaSharp;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            RasterizeDocument();
        }

        /// <summary>
        /// How to rasterize a document - save the document pages as images.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/rasterize-save-document-pages-as-picture-net-csharp-vb.php
        /// </remarks>
        static void RasterizeDocument()
        {
            // Rasterizing - it's process of converting the document pages into raster images.            
            // In this example we'll show how to rasterize/save a document page into PNG picture.

            string pngFile = @"Result.png";

            // Let's create a simple PDF document.
            DocumentCore dc = new DocumentCore();

            // Add new section.
            Section section = new Section(dc);
            dc.Sections.Add(section);

            // Let's set page size A4.
            section.PageSetup.PaperType = PaperType.A4;
            section.PageSetup.PageMargins.Left = LengthUnitConverter.Convert(10, LengthUnit.Millimeter, LengthUnit.Point);
            section.PageSetup.PageMargins.Right = LengthUnitConverter.Convert(10, LengthUnit.Millimeter, LengthUnit.Point);

            // Add any text on 1st page.
            Paragraph par1 = new Paragraph(dc);
            par1.ParagraphFormat.Alignment = HorizontalAlignment.Center;
            section.Blocks.Add(par1);

            // Let's create a characterformat for text in the 1st paragraph.
            CharacterFormat cf = new CharacterFormat() { FontName = "Verdana", Size = 86, FontColor = new SautinSoft.Document.Color(255, 255, 0) };

            Run text1 = new Run(dc, "You are welcome!");
            text1.CharacterFormat = cf;
            par1.Inlines.Add(text1);

            // Create the document paginator to get separate document pages.
            DocumentPaginator documentPaginator = dc.GetPaginator(new PaginatorOptions() { UpdateFields = true });

            // To get high-quality image, lets set 300 dpi.
            var DPI = new ImageSaveOptions();
            DPI.DpiX = 300;
            DPI.DpiY = 300;

            // Get the 1st page.
            DocumentPage page = documentPaginator.Pages[0];

            // Rasterize/convert the page into PNG image.
            page.Save(pngFile, DPI);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pngFile) { UseShellExecute = true });
        }
    }
}