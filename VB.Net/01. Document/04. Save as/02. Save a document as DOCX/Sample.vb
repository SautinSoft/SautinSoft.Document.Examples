Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        SaveToDocxFile()
        SaveToDocxStream()
    End Sub

    ''' <summary>
    ''' Creates a new document and saves it as DOCX file.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-docx-net-csharp-vb.php
    ''' </remarks>
    Sub SaveToDocxFile()
        ' Assume we already have a document 'dc'.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert("Hey from File!")

        Dim filePath As String = "Result-file.docx"

        dc.Save(filePath, New DocxSaveOptions())

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})

    End Sub

    ''' <summary>
    ''' Creates a new document and saves it as DOCX using MemoryStream.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-docx-net-csharp-vb.php
    ''' </remarks>
    Sub SaveToDocxStream()
        ' There variables are necessary only for demonstration purposes.
        Dim fileData() As Byte = Nothing
        Dim filePath As String = "Result-stream.docx"

        ' Assume we already have a document 'dc'.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert("Hey from MemoryStream!")

        ' Let's save our document to a MemoryStream.
        Using ms As New MemoryStream()
            dc.Save(ms, New DocxSaveOptions())
            fileData = ms.ToArray()
        End Using
        File.WriteAllBytes(filePath, fileData)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub
End Module