Imports System
Imports System.IO
Imports SautinSoft.Document
Imports System.Drawing

Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            RasterizeHtmlToPicture()
        End Sub

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

            Dim dpi As Integer = 300

            Dim pagesToRasterize As Integer = 2
            Dim currentPage As Integer = 1

            For Each page As DocumentPage In documentPaginator.Pages
                ' Save the page into Bitmap image with specified dpi and background.
                Dim picture As Bitmap = page.Rasterize(dpi, SautinSoft.Document.Color.White)

                ' Save the Bitmap to a PNG file.
                If currentPage = 1 Then
                    picture.Save(pngFile)
                ElseIf currentPage = 2 Then
                    ' Save the Bitmap to a JPEG file.
                    picture.Save(jpegFile, System.Drawing.Imaging.ImageFormat.Jpeg)
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
