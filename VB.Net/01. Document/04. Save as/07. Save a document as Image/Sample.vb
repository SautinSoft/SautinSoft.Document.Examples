Imports System
Imports System.IO
Imports System.Runtime
Imports SautinSoft.Document
Imports SkiaSharp


Module Sample
    Sub Main()
        SaveToImage()
    End Sub
    ''' Get your free 100-day key here:   
    ''' https://sautinsoft.com/start-for-free/
    ''' <summary>
    ''' Loads a document and saves all pages as images.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/save-document-as-image-net-csharp-vb.php
    ''' </remarks>
    Sub SaveToImage()
        Dim filePath As String = "..\..\..\example.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        Dim folderPath As String = Path.GetFullPath("Result-files")

        Dim dp As DocumentPaginator = dc.GetPaginator()

        For i As Integer = 0 To dp.Pages.Count - 1
            Dim page As DocumentPage = dp.Pages(i)
            ' For example, set DPI: 72, Background: White.

            Dim dpi As ImageSaveOptions = New ImageSaveOptions()
            dpi.DpiX = 72
            dpi.DpiY = 72
            Dim image As SKBitmap = page.Rasterize(DPI, SautinSoft.Document.Color.White)
            Directory.CreateDirectory(folderPath)
            image.Encode(New FileStream(folderPath + "\Page - " + i.ToString() + ".png", FileMode.Create), SkiaSharp.SKEncodedImageFormat.Png, 100)

        Next i
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(folderPath) With {.UseShellExecute = True})
    End Sub
End Module