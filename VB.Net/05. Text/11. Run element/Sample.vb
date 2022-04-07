Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        CalculateRuns()
    End Sub

    ''' <summary>
    ''' Loads an existing DOCX document and calculates all 'Run' objects.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/run-element-text-in-docx-document-net-csharp-vb.php
    ''' </remarks>
    Sub CalculateRuns()
        Dim filePath As String = "..\..\..\example.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        Dim filePathResult As String = "Result-file.docx"

        For Each par As Paragraph In dc.GetChildElements(True, ElementType.Paragraph)
            Dim totalRuns As Integer = par.GetChildElements(True, ElementType.Run).Count()

            Dim r As Run = New Run(dc, "<<This paragraph contains " & totalRuns.ToString() & " Run(s)>>", New CharacterFormat() With {
                    .BackgroundColor = Color.Yellow,
                    .Size = 10,
                    .FontColor = Color.Black
                })
            par.Content.End.Insert(r.Content)
        Next par
        dc.Save(filePathResult)
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePathResult) With {.UseShellExecute = True})
    End Sub
End Module