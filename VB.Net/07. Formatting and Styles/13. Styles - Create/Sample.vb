Imports SautinSoft.Document
Imports SautinSoft.Document.Tables

Namespace Sample
    Friend Class Sample
        Shared Sub Main(ByVal args() As String)
            CreateStyles()
        End Sub
        ''' <summary>
        ''' This sample shows how to create new styles and apply them to various elements.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/styles-create.php
        ''' </remarks>
        Public Shared Sub CreateStyles()
            Dim documentPath As String = "Styles.docx"

            ' Let's create document.
            Dim dc As New DocumentCore()

            ' Create 3 styles: CharacterStyle, ParagraphStyle and TableStyle

            ' 1. Create CharacterStyle
            Dim charStyle As New CharacterStyle("Green")
            charStyle.CharacterFormat = New CharacterFormat() With {
                .FontName = "Calibri",
                .Size = 20,
                .FontColor = Color.Green,
                .UnderlineStyle = UnderlineType.Single
            }

            ' 2. Create ParagraphStyle
            Dim parStyle As New ParagraphStyle("Center")
            parStyle.ParagraphFormat.Alignment = HorizontalAlignment.Center

            ' 3. Create TableStyle
            Dim tblStyle As New TableStyle("Blue Table")
            tblStyle.CellFormat.BackgroundColor = New Color("#358CCB")

            ' 4. Add the styles to the style collection.
            dc.Styles.Add(charStyle)
            dc.Styles.Add(parStyle)
            dc.Styles.Add(tblStyle)

            ' 5. Add some document content and apply the styles.
            ' Add a paragraph and text
            dc.Sections.Add(New Section(dc))
            Dim p As New Paragraph(dc)
            p.ParagraphFormat.Style = parStyle
            dc.Sections(0).Blocks.Add(p)

            p.Inlines.Add(New Run(dc, "Once upon a time, in a far away swamp, there lived an ogre named "))
            Dim r As New Run(dc, "Shrek")
            r.CharacterFormat.Style = charStyle
            p.Inlines.Add(r)
            p.Inlines.Add(New Run(dc, " whose precious solitude is suddenly shattered by an invasion of annoying fairy tale characters..."))

            ' Add a table (1 x 2).
            Dim tbl As New Table(dc, 1, 2)
            tbl.TableFormat.Style = tblStyle
            tbl.Rows(0).Cells(0).Content.Start.Insert("Column 1")
            tbl.Rows(0).Cells(1).Content.Start.Insert("Column 2")
            dc.Sections(0).Blocks.Add(tbl)

            ' Save our document into DOCX format.
            dc.Save(documentPath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
