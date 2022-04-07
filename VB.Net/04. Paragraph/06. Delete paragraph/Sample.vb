Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            DeleteParagraphs()
        End Sub
        ''' <summary>
        ''' Deletes a specific paragraphs in an existing DOCX and save it as new PDF.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/delete-paragraphs-in-docx-document-net-csharp-vb.php
        ''' </remarks>
        Private Shared Sub DeleteParagraphs()
            Dim filePath As String = "..\..\..\example.docx"
            Dim fileResult As String = "Result.pdf"

            Dim dc As DocumentCore = DocumentCore.Load(filePath)

            ' Note, remove paragraphs only inside the first section.
            Dim section As Section = dc.Sections(0)

            ' Let's remove all paragraphs containing the text "Jack".
            Dim i As Integer = 0
            Do While i < section.Blocks.Count
                If section.Blocks(i).Content.Find("Jack").Count() > 0 Then
                    section.Blocks.RemoveAt(i)
                    i -= 1
                End If
                i += 1
            Loop
            dc.Save(fileResult)
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(fileResult) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
