Imports System
Imports System.IO
Imports SautinSoft.Document
Imports System.Drawing

Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            RasterizePdfToPicture()
        End Sub

        ''' <summary>
        ''' Rasterizing - save PDF document as PNG and JPEG images.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/rasterize-save-pdf-document-as-png-jpeg-images-net-csharp-vb.php
        ''' </remarks>
        Private Shared Sub RasterizePdfToPicture()
            ' In this example we'll how rasterize/save 1st and 2nd pages of PDF document
            ' as PNG and JPEG images.
            Dim inputFile As String = "..\..\..\example.pdf"
            Dim jpegFile As String = "Result.jpg"
            Dim pngFile As String = "Result.png"

            Dim dc As DocumentCore = DocumentCore.Load(inputFile, New PdfLoadOptions() With {
                .DetectTables = False,
                .ConversionMode = PdfConversionMode.Exact
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
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(pngFile) With {.UseShellExecute = True})
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(jpegFile) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
