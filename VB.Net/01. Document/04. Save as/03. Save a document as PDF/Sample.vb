Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        SaveToPdfFile()
        SaveToPdfStream()
    End Sub

    ''' <summary>
    ''' Creates a new document and saves it as PDF file.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-pdf-net-csharp-vb.php
    ''' </remarks>
    Sub SaveToPdfFile()
        ' Assume we already have a document 'dc'.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert("Hey Guys and Girls!" & vbLf & "From file.", New CharacterFormat() With {
            .FontColor = Color.Green,
            .Size = 20
        })

        Dim filePath As String = "Result-file.pdf"

        dc.Save(filePath, New PdfSaveOptions() With {
                .Compliance = PdfCompliance.PDF_A1a,
                .PreserveFormFields = True
            })

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub

    ''' <summary>
    ''' Creates a new document and saves it as PDF/A using MemoryStream.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-pdf-net-csharp-vb.php
    ''' </remarks>
    Sub SaveToPdfStream()
        ' There variables are necessary only for demonstration purposes.
        Dim fileData() As Byte = Nothing
        Dim filePath As String = "Result-stream.pdf"

        ' Assume we already have a document 'dc'.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert("Hey Guys and Girls!" & vbLf & "From MemoryStream.", New CharacterFormat() With {
            .FontColor = Color.Orange,
            .Size = 20
        })

        ' Let's save our document to a MemoryStream.
        Using ms As New MemoryStream()
            dc.Save(ms, New PdfSaveOptions() With {
                .PageIndex = 0,
                .PageCount = 1,
                .Compliance = PdfCompliance.PDF_A1a
            })
            fileData = ms.ToArray()
        End Using
        File.WriteAllBytes(filePath, fileData)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})

    End Sub
End Module