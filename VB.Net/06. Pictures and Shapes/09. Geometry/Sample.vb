Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Module Sample
    Sub Main()
        Geometry()
    End Sub

    ''' <summary>
    ''' This sample shows how to work with shapes and geometry. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/geometry.php
    ''' </remarks>
    Sub Geometry()
        Dim pictPath As String = "..\..\..\image1.jpg"
        Dim documentPath As String = "Geometry.docx"

        ' Let's create a new document.
        Dim dc As New DocumentCore()

        ' Create shape 1 with preset geometry (Smiley Face).
        Dim shp1 As New Shape(dc, Layout.Floating(New HorizontalPosition(20.0F, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin), New VerticalPosition(80.0F, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), New Size(100, 100)))

        ' Specify outline and fill.
        shp1.Outline.Fill.SetSolid(New Color("358CCB"))
        shp1.Outline.Width = 3
        shp1.Fill.SetSolid(Color.Orange)

        ' Specify a figure.
        shp1.Geometry.SetPreset(Figure.SmileyFace)

        ' Create shape 2 with custom geometry path (using points array).
        Dim shp2 As New Shape(dc, Layout.Floating(New HorizontalPosition(85.0F, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin), New VerticalPosition(80.0F, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), New Size(100, 100)))

        ' Specify outline and fill using a picture.
        shp2.Outline.Fill.SetSolid(Color.Green)
        shp2.Outline.Width = 2

        ' Set the picture as fill for this shape.
        shp2.Fill.SetPicture(pictPath)

        ' Specify the maximum X and Y coordinates that should be used
        ' for within the path coordinate system.
        Dim size As New Size(1, 1)

        ' Specify the path points (draw a circle of 10 points).
        Dim points(9) As Point
        Dim a As Double = 0
        For i As Integer = 0 To 9
            points(i) = New Point(0.5 + Math.Sin(a) * 0.5, 0.5 + Math.Cos(a) * 0.5)
            a += 2 * Math.PI / 10
        Next i

        ' Create and add new custom path from specified points array.
        shp2.Geometry.SetCustom().AddPath(size, points, True)

        ' Create shape3 with custom geometry path (using path elements).
        Dim shp3 As New Shape(dc, Layout.Floating(New HorizontalPosition(150.0F, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin), New VerticalPosition(80.0F, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), New Size(100, 100)))

        ' Specify outline and fill.
        shp3.Outline.Fill.SetSolid(New Color(255, 0, 0))
        shp3.Outline.Width = 2
        shp3.Fill.SetSolid(Color.Yellow)

        ' Create and add new custom path.
        Dim path As CustomPath = shp3.Geometry.SetCustom().AddPath(New Size(1, 1))

        ' Specify path elements.
        path.MoveTo(New Point(0, 0))
        path.AddLine(New Point(0, 1))
        path.AddLine(New Point(1, 1))
        path.AddLine(New Point(1, 0))
        path.ClosePath()

        ' Add drawing elements to document.
        dc.Content.End.Insert(shp1.Content)
        dc.Content.End.Insert(shp2.Content)
        dc.Content.End.Insert(shp3.Content)

        ' Save the document to DOCX format.
        dc.Save(documentPath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub

End Module