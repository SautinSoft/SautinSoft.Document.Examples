Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Module Sample
    Sub Main()
        PictureAndShape()
    End Sub

    ''' <summary>
    ''' Creates a new document with shape containing a text and picture.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/pictures-and-shapes.php
    ''' </remarks>
    Sub PictureAndShape()
        Dim filePath As String = "Shape.docx"
        Dim imagePath As String = "..\..\..\image.jpg"

        Dim dc As New DocumentCore()

        ' 1. Shape with text.
        Dim shapeWithText As New Shape(dc, Layout.Floating(New HorizontalPosition(1, LengthUnit.Inch, HorizontalPositionAnchor.Page), New VerticalPosition(2, LengthUnit.Inch, VerticalPositionAnchor.Page), New Size(LengthUnitConverter.Convert(6, LengthUnit.Inch, LengthUnit.Point), LengthUnitConverter.Convert(1.5R, LengthUnit.Centimeter, LengthUnit.Point))))
        TryCast(shapeWithText.Layout, FloatingLayout).WrappingStyle = WrappingStyle.InFrontOfText
        shapeWithText.Text.Blocks.Add(New Paragraph(dc, New Run(dc, "This is the text in shape.", New CharacterFormat() With {.Size = 30})))
        shapeWithText.Outline.Fill.SetEmpty()
        shapeWithText.Fill.SetSolid(Color.Orange)
        dc.Content.End.Insert(shapeWithText.Content)

        ' 2. Picture with FloatingLayout:
        ' Floating layout means that the Picture (or Shape) is positioned by coordinates.
        Dim pic As New Picture(dc, imagePath)
        pic.Layout = FloatingLayout.Floating(New HorizontalPosition(50, LengthUnit.Millimeter, HorizontalPositionAnchor.Page), New VerticalPosition(20, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), New Size(LengthUnitConverter.Convert(10, LengthUnit.Centimeter, LengthUnit.Point), LengthUnitConverter.Convert(10, LengthUnit.Centimeter, LengthUnit.Point)))

        ' Set the wrapping style.
        TryCast(pic.Layout, FloatingLayout).WrappingStyle = WrappingStyle.BehindText

        ' Add our picture into the section.
        dc.Content.End.Insert(pic.Content)

        dc.Save(filePath)

        ' Show the result.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub
End Module