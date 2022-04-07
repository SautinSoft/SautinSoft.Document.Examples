Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        ConvertFromFile()
        ConvertFromStream()
    End Sub

    ''' <summary>
    ''' Convert PDF to DOCX (file to file).
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-document.php
    ''' </remarks>
    Sub ConvertFromFile()
        Dim inpFile As String = "..\..\..\example.pdf"
        Dim outFile As String = "Result.docx"

        Dim dc As DocumentCore = DocumentCore.Load(inpFile)
        dc.Save(outFile)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
    End Sub

    ''' <summary>
    ''' Convert PDF to HTML (using Stream).
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-document.php
    ''' </remarks>
    Sub ConvertFromStream()

        ' We need files only for demonstration purposes.
        ' The conversion process will be done completely in memory.
        Dim inpFile As String = "..\..\..\example.pdf"
        Dim outFile As String = "Result.html"
        Dim inpData() As Byte = File.ReadAllBytes(inpFile)
        Dim outData() As Byte = Nothing

        Using msInp As New MemoryStream(inpData)

            ' Load a document.
            Dim dc As DocumentCore = DocumentCore.Load(msInp, New PdfLoadOptions() With {
                .PreserveGraphics = True,
                .DetectTables = True
            })

            ' Save the document to HTML-fixed format.
            Using outMs As New MemoryStream()
                dc.Save(outMs, New HtmlFixedSaveOptions() With {
                    .CssExportMode = CssExportMode.Inline,
                    .EmbedImages = True
                })
                outData = outMs.ToArray()
            End Using
            ' Show the result for demonstration purposes.
            If outData IsNot Nothing Then
                File.WriteAllBytes(outFile, outData)
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
            End If
        End Using
    End Sub
End Module