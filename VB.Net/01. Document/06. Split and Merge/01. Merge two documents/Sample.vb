Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        MergeTwoDocuments()
    End Sub

    ''' <summary>
    ''' How to merge two documents: DOCX and PDF.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/split-and-merge-content-net-csharp-vb.php
    ''' </remarks>
    Sub MergeTwoDocuments()
        ' Path to our combined document.
        Dim singleFilePath As String = "Single.docx"

        Dim supportedFiles() As String = {"..\..\..\example.docx", "..\..\..\example.pdf"}

        ' Create single document.
        Dim dcSingle As New DocumentCore()

        For Each file As String In supportedFiles
            Dim dc As DocumentCore = DocumentCore.Load(file)

            Console.WriteLine("Adding: {0}...", Path.GetFileName(file))

            ' Create import session.
            Dim session As New ImportSession(dc, dcSingle, StyleImportingMode.KeepSourceFormatting)

            ' Loop through all sections in the source document.
            For Each sourceSection As Section In dc.Sections
                ' Because we are copying a section from one document to another,
                ' it is required to import the Section into the destination document.
                ' This adjusts any document-specific references to styles, bookmarks, etc.
                '
                ' Importing a element creates a copy of the original element, but the copy
                ' is ready to be inserted into the destination document.
                Dim importedSection As Section = dcSingle.Import(Of Section)(sourceSection, True, session)

                ' First section start from new page.
                If dc.Sections.IndexOf(sourceSection) = 0 Then
                    importedSection.PageSetup.SectionStart = SectionStart.NewPage
                End If

                ' Now the new section can be appended to the destination document.
                dcSingle.Sections.Add(importedSection)
            Next sourceSection
        Next file

        ' Save single document to a file.
        dcSingle.Save(singleFilePath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(singleFilePath) With {.UseShellExecute = True})
    End Sub
End Module