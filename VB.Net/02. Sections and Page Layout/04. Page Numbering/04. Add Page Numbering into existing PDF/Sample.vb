Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        AddPageNumbering()
    End Sub

    ''' <summary>
    ''' Add page numbering into an existing PDF document.
    ''' </summary>
    ''' <remarks>
    ''' https://www.sautinsoft.com/products/document/help/net/developer-guide/add-page-numbering-in-pdf-net-csharp-vb.php
    ''' </remarks>
    Sub AddPageNumbering()
        Dim inpFile As String = "..\..\..\shrek.pdf"
        Dim outFile As String = "With Pages.pdf"
        Dim dc As DocumentCore = DocumentCore.Load(inpFile)

        ' We place our page numbers into the footer.
        ' Therefore we've to create a footer.
        Dim footer As New HeaderFooter(dc, HeaderFooterType.FooterDefault)

        ' Create a new paragraph to insert a page numbering.
        ' So that, our page numbering looks as: Page N of M.
        Dim par As New Paragraph(dc)
        par.ParagraphFormat.Alignment = HorizontalAlignment.Right
        Dim cf As New CharacterFormat() With {
                .FontName = "Consolas",
                .Size = 14.0
            }
        par.Content.Start.Insert("Page ", cf.Clone())

        ' Page numbering is a Field.
        ' Create two fields: FieldType.Page and FieldType.NumPages.
        Dim fPage As New Field(dc, FieldType.Page)
        fPage.CharacterFormat = cf.Clone()
        par.Content.End.Insert(fPage.Content)
        par.Content.End.Insert(" of ", cf.Clone())
        Dim fPages As New Field(dc, FieldType.NumPages)
        fPages.CharacterFormat = cf.Clone()
        par.Content.End.Insert(fPages.Content)
        footer.Blocks.Add(par)

        For Each s As Section In dc.Sections
            s.HeadersFooters.Add(footer.Clone(True))
        Next s
        dc.Save(outFile)

        ' Open the results for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(inpFile) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
    End Sub
End Module