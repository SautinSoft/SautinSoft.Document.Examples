Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        DeletePageNumbering()
    End Sub

    ''' <summary>
    ''' How to delete the page numbering in an existing DOCX document. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/remove-page-numbering-in-docx-document-net-csharp-vb.php
    ''' </remarks>
    Sub DeletePageNumbering()
        Dim inpFile As String = "..\..\..\PageNumbering.docx"
        ' Save the output document into PDF format,
        ' you may change it to any from: DOCX, HTML, PDF, RTF.
        Dim outFile As String = "PageNumbering - deleted.pdf"

        ' Load a document from DOCX file (PageNumbering.docx).
        ' We've created PageNumbering.docx in the previous example.
        Dim dc As DocumentCore = DocumentCore.Load(inpFile)
        Dim s As Section = dc.Sections(0)

        For Each hf As HeaderFooter In s.HeadersFooters
            For Each field As Field In hf.GetChildElements(True, ElementType.Field).Reverse()
                ' Page numbering is a Field,
                ' so we have to find the fields with the type Page or NumPages and remove.
                If field.FieldType = FieldType.Page OrElse field.FieldType = FieldType.NumPages Then
                    ' Also assume that our page numbering located in a paragraph,
                    ' so let's remove the paragraph's content too.
                    If TypeOf field.Parent Is Paragraph Then
                        TryCast(field.Parent, Paragraph).Inlines.Clear()
                    End If

                    ' If we'll delete only the fields (field.Content.Delete()), our paragraph
                    ' may contain text "Page of ".
                    ' Based on this, we remove the whole paragraph.
                End If
            Next field
        Next hf
        dc.Save(outFile)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(inpFile) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
    End Sub
End Module