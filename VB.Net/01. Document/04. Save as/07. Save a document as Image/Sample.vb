Imports System
Imports System.IO
Imports SautinSoft.Document
Imports System.Drawing
Imports System.Drawing.Imaging

Module Sample
    Sub Main()
        SaveToImage()
    End Sub

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
            Dim image As Bitmap = page.Rasterize(72, SautinSoft.Document.Color.White)
            Directory.CreateDirectory(folderPath)
            image.Save(folderPath & "\Page - " & i.ToString() & ".png", ImageFormat.Png)
        Next i
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(folderPath) With {.UseShellExecute = True})
    End Sub
End Module