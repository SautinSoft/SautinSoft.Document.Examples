Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Module Sample
    Sub Main()
        InsertTextAsWatermark()
    End Sub

    ''' <summary>
    ''' How to insert a text watermark in the existing PDF, DOCX, any document to the each page.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-text-watermark-in-the-existing-pdf-docx-document-to-each-page-net-csharp-vb.php
    ''' </remarks>
    Sub InsertTextAsWatermark()
        ' Let us say, we want to insert a textual "Watermark" at the each page of the
        ' DOCX document and specify angle of 45 degree for it.

        ' If we'll insert the text at the document header (or footer), so it will appear on the each page.

        ' Also let's insert our Watermark behind the main content.
        Dim inpFile As String = "..\..\..\example.docx"
        Dim outFile As String = "Result.docx"

        ' 1. Load an existing DOCX document.
        Dim dc As DocumentCore = DocumentCore.Load(inpFile)

        ' 2. Create a Shape with our Watermark text.
        ' Place the watermark:
        ' 30mm - from the page left;            
        ' 150mm - from the page top.
        ' 60 - angle.
        Dim posFromLeft As Single = 30.0F
        Dim posFromTop As Single = 150.0F
        Dim angle As Single = -60.0F
        ' Size of the Shape, 200mm x 40mm.            
        Dim size As SautinSoft.Document.Drawing.Size = New Size(200.0F, 40.0F, LengthUnit.Millimeter)
        Dim watermark As New Shape(dc, New FloatingLayout(New HorizontalPosition(posFromLeft, LengthUnit.Millimeter, HorizontalPositionAnchor.Page), New VerticalPosition(posFromTop, LengthUnit.Millimeter, VerticalPositionAnchor.Page), size))
        ' Rotate shape.
        watermark.Rotation = angle
        ' Create the text.
        Dim text As New Run(dc, "Watermark!", New CharacterFormat() With {
                .Size = 100.0F,
                .FontColor = Color.Black,
                .FontName = "Arial"
            })
        watermark.Text.Blocks.Add(New Paragraph(dc, text))
        ' Set shape Behind the main document contents.
        TryCast(watermark.Layout, FloatingLayout).WrappingStyle = WrappingStyle.BehindText
        ' Remove the shape borders.
        watermark.Outline.Fill.SetEmpty()

        ' 3. Iterate through Sections, and insert our Watermark to the default header of the each section.
        For Each section As Section In dc.Sections
            ' 2.1. Check the document header, maybe is it already exist?
            Dim header = dc.Sections(0).HeadersFooters(HeaderFooterType.HeaderDefault)

            If header Is Nothing Then
                ' Create a new header, add it into the section.
                header = New HeaderFooter(dc, HeaderFooterType.HeaderDefault)
                ' Add the header to the section.
                section.HeadersFooters.Add(header)
            End If
            ' Add the watermark to the header.
            header.Content.End.Insert(watermark.Content)
        Next section

        ' 4. Save the document back.
        dc.Save(outFile)
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
    End Sub

End Module