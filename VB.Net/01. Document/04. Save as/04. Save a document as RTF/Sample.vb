Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        SaveToRtfFile()
        SaveToRtfStream()
    End Sub

    ''' <summary>
    ''' Creates a new document and saves it as RTF file.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-rtf-net-csharp-vb.php
    ''' </remarks>
    Sub SaveToRtfFile()
        ' Assume we already have a document 'dc'.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert("Hey Guys and Girls!" & vbLf & "From file.", New CharacterFormat() With {
            .FontColor = Color.Green,
            .Size = 20
        })

        Dim filePath As String = "Result-file.rtf"

        dc.Save(filePath, New RtfSaveOptions())

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub

    ''' <summary>
    ''' Creates a new document and saves it as RTF using MemoryStream.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-rtf-net-csharp-vb.php
    ''' </remarks>
    Sub SaveToRtfStream()
        ' There variables are necessary only for demonstration purposes.
        Dim fileData() As Byte = Nothing
        Dim filePath As String = "Result-stream.rtf"

        ' Assume we already have a document 'dc'.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert("Hey Guys and Girls!" & vbLf & "From MemoryStream.", New CharacterFormat() With {
            .FontColor = Color.Orange,
            .Size = 20
        })

        ' Let's save our document to a MemoryStream.
        Using ms As New MemoryStream()
            dc.Save(ms, New RtfSaveOptions())
            fileData = ms.ToArray()
        End Using
        File.WriteAllBytes(filePath, fileData)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub
End Module