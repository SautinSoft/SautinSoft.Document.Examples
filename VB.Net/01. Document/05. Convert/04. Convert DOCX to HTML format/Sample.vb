Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        ConvertFromFile()
        ConvertFromStream()
    End Sub

    ''' <summary>
    ''' Convert DOCX to HTML (file to file).
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-docx-to-html-in-csharp-vb.php
    ''' </remarks>
    Sub ConvertFromFile()
        Dim inpFile As String = "..\..\..\example.docx"
        Dim outFile As String = "Result.html"

        Dim dc As DocumentCore = DocumentCore.Load(inpFile)
        dc.Save(outFile)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
    End Sub

    ''' <summary>
    ''' Convert DOCX to HTML (using Stream).
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-docx-to-html-in-csharp-vb.php
    ''' </remarks>
    Sub ConvertFromStream()

        ' We need files only for demonstration purposes.
        ' The conversion process will be done completely in memory.
        Dim inpFile As String = "..\..\..\example.docx"
        Dim outFile As String = "ResultStream.html"
        Dim inpData() As Byte = File.ReadAllBytes(inpFile)
        Dim outData() As Byte = Nothing

        Using msInp As New MemoryStream(inpData)

            ' Load a document.
            Dim dc As DocumentCore = DocumentCore.Load(msInp, New DocxLoadOptions())

            ' Save the document to HTML format.
            Using outMs As New MemoryStream()
                dc.Save(outMs, New HtmlFixedSaveOptions())
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