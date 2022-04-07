Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            InsertParagraph()
        End Sub
        ''' <summary>
        ''' Inserts a new paragraph into an existing PDF document.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-paragraphs-to-pdf-document-net-csharp-vb.php
        ''' </remarks>
        Private Shared Sub InsertParagraph()
            Dim inpFile As String = "..\..\..\example.pdf"
            Dim outFile As String = "Result.pdf"

            Dim dc As DocumentCore = DocumentCore.Load(inpFile)
            Dim p As New Paragraph(dc)
            p.Content.Start.Insert("Alexander Pushkin was a great russian romantic poet " & "and writer who is considered by a lot of people as the best russian poet and the founder " & "of contemporary russian literature.", New CharacterFormat() With {
                .Size = 20,
                .FontName = "Verdana",
                .FontColor = New Color("#358CCB")
            })
            p.ParagraphFormat.Alignment = HorizontalAlignment.Justify

            ' Insert the paragraph as 1st element in the 1st section.
            dc.Sections(0).Blocks.Insert(0, p)

            dc.Save(outFile)
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})

        End Sub
    End Class
End Namespace
