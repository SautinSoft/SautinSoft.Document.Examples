Imports System
Imports System.IO
Imports System.Collections.Generic
Imports SautinSoft.Document

Namespace Sample
    Friend Class Sample

        Shared Sub Main(ByVal args() As String)
            MergeMultipleDocuments()
        End Sub

        ''' <summary>
        ''' This sample shows how to merge multiple DOCX, RTF, PDF and Text files.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/merge-multiple-files-net-csharp-vb.php
        ''' </remarks>
        Public Shared Sub MergeMultipleDocuments()
            ' Path to our combined document.
            Dim singlePDFPath As String = "Single.pdf"
            Dim workingDir As String = "..\..\..\"

            Dim supportedFiles As New List(Of String)()
            ' Fill the collection 'supportedFiles' by *.docx, *.pdf and *.txt.
            For Each file As String In Directory.GetFiles(workingDir, "*.*")
                Dim ext As String = Path.GetExtension(file).ToLower()

                If ext = ".docx" OrElse ext = ".pdf" OrElse ext = ".txt" Then
                    supportedFiles.Add(file)
                End If
            Next file


            ' Create single pdf.
            Dim singlePDF As New DocumentCore()

            For Each file As String In supportedFiles
                Dim dc As DocumentCore = DocumentCore.Load(file)

                Console.WriteLine("Adding: {0}...", Path.GetFileName(file))

                ' Create import session.
                Dim session As New ImportSession(dc, singlePDF, StyleImportingMode.KeepSourceFormatting)

                ' Loop through all sections in the source document.
                For Each sourceSection As Section In dc.Sections
                    ' Because we are copying a section from one document to another,
                    ' it is required to import the Section into the destination document.
                    ' This adjusts any document-specific references to styles, bookmarks, etc.
                    '
                    ' Importing a element creates a copy of the original element, but the copy
                    ' is ready to be inserted into the destination document.
                    Dim importedSection As Section = singlePDF.Import(Of Section)(sourceSection, True, session)

                    ' First section start from new page.
                    If dc.Sections.IndexOf(sourceSection) = 0 Then
                        importedSection.PageSetup.SectionStart = SectionStart.NewPage
                    End If

                    ' Now the new section can be appended to the destination document.
                    singlePDF.Sections.Add(importedSection)
                Next sourceSection
            Next file

            ' Save single PDF to a file.
            singlePDF.Save(singlePDFPath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(singlePDFPath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
