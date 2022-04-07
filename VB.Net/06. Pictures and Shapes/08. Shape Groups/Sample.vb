Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Module Sample
    Sub Main()
        ShapeGroups()
    End Sub

    ''' <summary>
    ''' This sample shows how to work with shape groups. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/shape-groups.php
    ''' </remarks>
    Sub ShapeGroups()
        Dim pictPath As String = "..\..\..\image1.jpg"
        Dim documentPath As String = "ShapeGroups.docx"

        ' Let's create a document.
        Dim dc As New DocumentCore()

        ' Create floating layout.
        Dim hp As New HorizontalPosition(HorizontalPositionType.Center, HorizontalPositionAnchor.Page)
        Dim vp As New VerticalPosition(5.0F, LengthUnit.Centimeter, VerticalPositionAnchor.TopMargin)
        Dim fl As New FloatingLayout(hp, vp, New Size(300, 300))

        ' Create group.
        Dim group As New ShapeGroup(dc, fl)

        ' Specify the size dimensions of the child extents rectangle.
        group.ChildSize = New Size(100, 100)

        ' Create a child shape#1 (inside group) with preset geometry.
        ' Specify shape's size and offset relative to group's ChildSize (100x100).
        Dim shape1 As New Shape(dc, New GroupLayout(New Point(0, 0), New Size(50, 50)))

        ' Specify outline and fill.
        shape1.Outline.Fill.SetSolid(New Color("#358CCB"))
        shape1.Outline.Width = 2
        shape1.Fill.SetSolid(Color.Orange)

        ' Shape will be rectangle.
        shape1.Geometry.SetPreset(Figure.Rectangle)

        ' Create picture and add it into the group.
        Dim picture As New Picture(dc, Layout.Group(New Point(50, 50), New Size(50, 50)), pictPath)

        ' Specify picture fill mode.
        picture.ImageData.FillMode = PictureFillMode.Stretch

        ' Add shape and picture into our group.
        group.ChildShapes.Add(shape1)
        group.ChildShapes.Add(picture)

        ' Add our group into the document.
        dc.Content.End.Insert(group.Content)

        ' Save our document into DOCX format.
        dc.Save(documentPath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub

End Module