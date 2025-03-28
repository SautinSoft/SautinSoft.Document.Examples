Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SkiaSharp

Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            RasterizeHtmlToPicture()
        End Sub
        ''' Get your free trial key here:   
        ''' https://sautinsoft.com/start-for-free/
        ''' <summary>
        ''' Rasterizing - save HTML document as PNG and JPEG images.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/rasterize-save-html-document-as-png-jpeg-images-net-csharp-vb.php
        ''' </remarks>
        Private Shared Sub RasterizeHtmlToPicture()
            ' In this example we'll how rasterize/save 1st and 2nd pages of HTML document
            ' as PNG and JPEG images.
            Dim inputFile As String = "..\..\..\example.html"
            Dim jpegFile As String = "Result.jpg"
            Dim pngFile As String = "Result.png"
            ' The file format is detected automatically from the file extension: ".html".
            ' But as shown in the example below, we can specify HtmlLoadOptions as 2nd parameter
            ' to explicitly set that a loadable document has HTML format.
            Dim dc As DocumentCore = DocumentCore.Load(inputFile, New HtmlLoadOptions() With {
                .PageSetup = New PageSetup() With {
                    .PaperType = PaperType.A5,
                    .Orientation = Orientation.Landscape
                }
            })

            Dim documentPaginator As DocumentPaginator = dc.GetPaginator(New PaginatorOptions() With {.UpdateFields = True})

            Dim pagesToRasterize As Integer = 2
            Dim currentPage As Integer = 1

            For Each page As DocumentPage In documentPaginator.Pages

                ' Save to a PNG file.
                If currentPage = 1 Then
                    page.Save(pngFile, New ImageSaveOptions() With {
                        .Format = ImageSaveFormat.Png, .DpiX = 72, .DpiY = 72})

                    ' Save to a JPEG file.
                ElseIf currentPage = 2 Then
                    page.Save(jpegFile, New ImageSaveOptions() With {
                        .Format = ImageSaveFormat.Jpeg, .DpiX = 72, .DpiY = 72})
                End If

                currentPage += 1

                If currentPage > pagesToRasterize Then
                    Exit For
                End If
            Next page

            ' Open the results for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(jpegFile) With {.UseShellExecute = True})
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(pngFile) With {.UseShellExecute = True})

        End Sub
    End Class
End Namespace
