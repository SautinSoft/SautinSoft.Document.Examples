Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        SaveToFile()
        SaveToStream()
    End Sub

    ''' <summary>
    ''' Creates a new document and saves it as DOCX file.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document.php
    ''' </remarks>
    Sub SaveToFile()
        ' Assume we already have a document 'dc'.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert("Hey Guys and Girls!")

        Dim filePath As String = "Result.docx"

        ' The file format will be detected automatically from the file extension: ".docx".
        dc.Save(filePath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub

    ''' <summary>
    ''' Creates a new document and saves it as PDF using MemoryStream.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document.php
    ''' </remarks>
    Sub SaveToStream()
        ' There variables are necessary only for demonstration purposes.
        Dim fileData() As Byte = Nothing
        Dim filePath As String = "Result.pdf"

        ' Assume we already have a document 'dc'.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert("Hey Guys and Girls!")

        ' Let's save our document to a MemoryStream.
        Using ms As New MemoryStream()
            ' 2nd parameter: we've explicitly set to save our document in PDF format.
            dc.Save(ms, New PdfSaveOptions())

            fileData = ms.ToArray()
        End Using

        File.WriteAllBytes(filePath, fileData)
        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub
End Module