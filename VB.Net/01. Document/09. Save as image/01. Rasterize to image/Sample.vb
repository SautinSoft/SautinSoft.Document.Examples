Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SkiaSharp

Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            RasterizeDocument()
        End Sub
        ''' Get your free 30-day key here:   
        ''' https://sautinsoft.com/start-for-free/
        ''' <summary>
        ''' How to rasterize a document - save the document pages as images.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/rasterize-save-document-pages-as-picture-net-csharp-vb.php
        ''' </remarks>
        Private Shared Sub RasterizeDocument()
            ' Rasterizing - it's process of converting the document pages into raster images.            
            ' In this example we'll show how to rasterize/save a document page into PNG picture.

            Dim pngFile As String = "Result.png"

            ' Let's create a simple PDF document.
            Dim dc As New DocumentCore()

            ' Add new section.
            Dim section As New Section(dc)
            dc.Sections.Add(section)

            ' Let's set page size A4.
            section.PageSetup.PaperType = PaperType.A4
            section.PageSetup.PageMargins.Left = LengthUnitConverter.Convert(10, LengthUnit.Millimeter, LengthUnit.Point)
            section.PageSetup.PageMargins.Right = LengthUnitConverter.Convert(10, LengthUnit.Millimeter, LengthUnit.Point)

            ' Add any text on 1st page.
            Dim par1 As New Paragraph(dc)
            par1.ParagraphFormat.Alignment = HorizontalAlignment.Center
            section.Blocks.Add(par1)

            ' Let's create a characterformat for text in the 1st paragraph.
            Dim cf As New CharacterFormat() With {
                .FontName = "Verdana",
                .Size = 86,
                .FontColor = New SautinSoft.Document.Color("#FFFF00")
            }

            Dim text1 As New Run(dc, "You are welcome!")
            text1.CharacterFormat = cf
            par1.Inlines.Add(text1)

            ' Create the document paginator to get separate document pages.
            Dim documentPaginator As DocumentPaginator = dc.GetPaginator(New PaginatorOptions() With {.UpdateFields = True})

            ' To get high-quality image, lets set 300 dpi.
            Dim dpi As ImageSaveOptions = New ImageSaveOptions()
            dpi.DpiX = 300
            dpi.DpiY = 300

            ' Get the 1st page.
            Dim page As DocumentPage = documentPaginator.Pages(0)

            ' Rasterize/convert the page into PNG image.
            Dim image As SKBitmap = page.Rasterize(dpi, SautinSoft.Document.Color.LightGray)

            image.Encode(New FileStream(pngFile, FileMode.Create), SkiaSharp.SKEncodedImageFormat.Png, 100)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(pngFile) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
