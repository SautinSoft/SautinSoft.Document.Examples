Imports System
Imports System.IO
Imports System.Collections.Generic
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Module Sample
    Sub Main()
        ExtractPictures()
    End Sub

    ''' <summary>
    ''' Extract all pictures from document (PDF, DOCX, RTF, HTML).
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/extract-pictures.php
    ''' </remarks>
    Sub ExtractPictures()
        ' Path to a document where to extract pictures.
        Dim filePath As String = "..\..\..\example.pdf"

        ' Directory to store extracted pictures:
        Dim imgDir As New DirectoryInfo("Extracted Pictures")
        imgDir.Create()
        Dim imgTemplateName As String = "Picture"

        ' Here we store extracted images.
        Dim imgInventory As New List(Of ImageData)()

        ' Load the document.
        Dim dc As DocumentCore = DocumentCore.Load(filePath)

        ' Extract all images from document, skip duplicates.
        For Each pict As Picture In dc.GetChildElements(True, ElementType.Picture)
            ' Let's avoid the adding of duplicates.
            If imgInventory.Exists((Function(img) (img.GetStream().Length = pict.ImageData.GetStream().Length))) = False Then
                imgInventory.Add(pict.ImageData)
            End If
        Next pict

        ' Save all images.
        For i As Integer = 0 To imgInventory.Count - 1
            Dim imagePath As String = Path.Combine(imgDir.FullName, String.Format("{0}{1}.{2}", imgTemplateName, i + 1, imgInventory(i).Format.ToString().ToLower()))
            File.WriteAllBytes(imagePath, imgInventory(i).GetStream().ToArray())
        Next i

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(imgDir.FullName) With {.UseShellExecute = True})
    End Sub
End Module