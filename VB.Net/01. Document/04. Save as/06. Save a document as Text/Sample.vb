Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        SaveToTextFile()
        SaveToTextStream()
    End Sub

    ''' <summary>
    ''' Creates a new document and saves it as Text file.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-text-net-csharp-vb.php
    ''' </remarks>
    Sub SaveToTextFile()
        ' Assume we already have a document 'dc'.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert("Hey Guys and Girls!")

        Dim filePath As String = "Result-file.txt"

        dc.Save(filePath, New TxtSaveOptions() With {
            .Encoding = System.Text.Encoding.UTF8,
            .ParagraphBreak = Environment.NewLine
        })

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub

    ''' <summary>
    ''' Creates a new document and saves it as Text using MemoryStream.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-text-net-csharp-vb.php
    ''' </remarks>
    Sub SaveToTextStream()
        ' There variables are necessary only for demonstration purposes.
        Dim fileData() As Byte = Nothing
        Dim filePath As String = "Result-stream.txt"

        ' Assume we already have a document 'dc'.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert("Hey Guys and Girls!")

        ' Let's save our document to a MemoryStream.
        Using ms As New MemoryStream()
            dc.Save(ms, New TxtSaveOptions())
            fileData = ms.ToArray()
        End Using
        File.WriteAllBytes(filePath, fileData)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub
End Module