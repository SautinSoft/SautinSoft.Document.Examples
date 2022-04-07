Option Infer On

Imports System
Imports SautinSoft.Document
Imports System.IO
Imports System.Linq
Imports System.Text
Module Sample
    Sub Main()
        ImportingElement()
    End Sub

    ''' <summary>
    ''' Copy a document to other document. Supported any formats (PDF, DOCX, RTF, HTML).
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/importing-element-net-csharp-vb.php
    ''' </remarks>
    Sub ImportingElement()
        Dim file1 As String = "..\..\..\digitalsignature.docx"
        Dim file2 As String = "..\..\..\Parsing.docx"
        Dim resultFile As String = "Importing.docx"

        ' Load files.
        Dim dc As DocumentCore = DocumentCore.Load(file1)
        Dim dc1 As DocumentCore = DocumentCore.Load(file2)

        ' New Import Session to improve performance.
        Dim session = New ImportSession(dc1, dc)

        ' Import all sections from source document.
        For Each sourceSection As Section In dc1.Sections
            Dim destinationSection As Section = dc.Import(sourceSection, True, session)
            dc.Sections.Add(destinationSection)
        Next sourceSection

        ' Save the result.
        dc.Save(resultFile)

        ' Show the result.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultFile) With {.UseShellExecute = True})
    End Sub
End Module