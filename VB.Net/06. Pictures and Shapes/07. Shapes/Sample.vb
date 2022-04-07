Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Module Sample
    Sub Main()
        Shapes()
    End Sub

    ''' <summary>
    ''' This sample shows how to work with shapes. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/shapes.php
    ''' </remarks>
    Sub Shapes()
        Dim documentPath As String = "Shapes.docx"

        ' Let's create a new document.
        Dim dc As New DocumentCore()

        ' Create shape 1 with fill and outline.
        Dim shp1 As New Shape(dc, Layout.Floating(New HorizontalPosition(25.0F, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin), New VerticalPosition(20.0F, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), New Size(200, 100)))

        ' Specify outline and fill using a picture.
        shp1.Outline.Fill.SetSolid(Color.DarkGreen)
        shp1.Outline.Width = 2

        ' Set fill for this shape.
        shp1.Fill.SetSolid(Color.Orange)

        ' Create shape 2 with some text inside, 100mm*20mm.
        Dim shp2 As New Shape(dc, Layout.Floating(New HorizontalPosition(100.0F, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin), New VerticalPosition(20.0F, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), New Size(LengthUnitConverter.Convert(100.0F, LengthUnit.Millimeter, LengthUnit.Point), LengthUnitConverter.Convert(20.0F, LengthUnit.Millimeter, LengthUnit.Point))))

        ' Specify outline and fill using a picture.
        shp2.Outline.Fill.SetSolid(Color.LightGray)
        shp2.Outline.Width = 0.5

        ' Create a new paragraph with a formatted text.
        Dim p As New Paragraph(dc)
        Dim run1 As New Run(dc, "Welcome to International Software Developer conference!")
        run1.CharacterFormat.FontName = "Helvetica"
        run1.CharacterFormat.Size = 14.0F
        run1.CharacterFormat.Italic = True
        p.Inlines.Add(run1)

        ' Add the paragraph into the shp2.Text property.
        shp2.Text.Blocks.Add(p)

        ' Add our shapes into the document.
        dc.Content.End.Insert(shp1.Content)
        dc.Content.End.Insert(shp2.Content)

        ' Save the document to DOCX format.
        dc.Save(documentPath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub
End Module