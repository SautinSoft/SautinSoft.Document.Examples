Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        EditPageNumbering()
    End Sub

    ''' <summary>
    ''' How to edit Page Numbering in an existing DOCX document. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/edit-page-numbering-in-docx-document-net-csharp-vb.php
    ''' </remarks>
    Sub EditPageNumbering()
        ' Modify the Page Numbering from: "Page N of M" to "Página N".
        Dim inpFile As String = "..\..\..\PageNumbering.docx"
        Dim outFile As String = "PageNumbering-modified.pdf"

        ' Load a DOCX document with Page Numbering (PageNumbering.docx).
        Dim dc As DocumentCore = DocumentCore.Load(inpFile)
        Dim s As Section = dc.Sections(0)

        For Each hf As HeaderFooter In s.HeadersFooters
            For Each field As Field In hf.GetChildElements(True, ElementType.Field).Reverse()
                ' Page numbering is a Field,
                ' so we have to find the fields with the type Page or NumPages and remove.
                If field.FieldType = FieldType.Page OrElse field.FieldType = FieldType.NumPages Then
                    ' Save the paragraph containing the page numbering and character formatting.
                    Dim parWithNumbering As Paragraph = Nothing
                    Dim cf As CharacterFormat = field.CharacterFormat.Clone()

                    ' Also assume that our page numbering located in a paragraph,
                    ' so let's remove the paragraph's content too.
                    If TypeOf field.Parent Is Paragraph Then
                        parWithNumbering = (TryCast(field.Parent, Paragraph))
                        ' If we'll delete only the fields (field.Content.Delete()), our paragraph
                        ' may contain text "Page of ".
                        ' Based on this, we remove all items in this paragraph.
                        ' Fields {Page} and {NumPages} will be also deleted here, because they are Inlines too.
                        parWithNumbering.Inlines.Clear()
                    End If

                    If parWithNumbering IsNot Nothing Then
                        ' Insert new Page Numbering: "Página N":                        
                        ' Insert "Página ".
                        parWithNumbering.Inlines.Add(New Run(dc, "Página ", cf.Clone()))
                        ' Insert field {Page}.
                        Dim fPage As New Field(dc, FieldType.Page)
                        fPage.CharacterFormat = cf.Clone()
                        parWithNumbering.Inlines.Add(fPage)
                    End If
                End If
            Next field
        Next hf
        dc.Save(outFile)

        ' Open the results for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(inpFile) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
    End Sub
End Module