Imports System
Imports SautinSoft.Document
Imports System.IO
Imports System.Linq
Imports System.Text

Module Sample
    Sub Main()
        CloningElement()
    End Sub

    ''' <summary>
    ''' How to clone an element in DOCX document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/cloning-element-net-csharp-vb.php
    ''' </remarks>
    Sub CloningElement()
        Dim filePath As String = "..\..\..\Parsing.docx"
        Dim cloningFile As String = "Cloning.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)

        ' Clone section.
        dc.Sections.Add(dc.Sections(0).Clone(True))

        ' Clone paragraphs.
        For Each item As Block In dc.Sections(0).Blocks
            dc.Sections.Last().Blocks.Add(item.Clone(True))
        Next item

        ' Save the result.
        dc.Save(cloningFile)

        ' Show the results.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(cloningFile) With {.UseShellExecute = True})
    End Sub
End Module