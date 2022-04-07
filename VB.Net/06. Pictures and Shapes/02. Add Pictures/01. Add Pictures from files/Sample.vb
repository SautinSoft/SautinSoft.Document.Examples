Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Module Sample
    Sub Main()
        AddPictures()
    End Sub

    ''' <summary>
    ''' How to add pictures into a document. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/add-pictures.php
    ''' </remarks>
    Sub AddPictures()
        Dim documentPath As String = "Pictures.docx"
        Dim pictPath As String = "..\..\..\image1.jpg"

        ' Let's create a simple document.
        Dim dc As New DocumentCore()

        ' Add a new section.
        Dim s As New Section(dc)
        dc.Sections.Add(s)

        ' 1. Picture with InlineLayout:

        ' Create a new paragraph with picture.
        Dim par As New Paragraph(dc)
        s.Blocks.Add(par)
        par.ParagraphFormat.Alignment = HorizontalAlignment.Left

        ' Add some text content.
        par.Content.End.Insert("Shrek and Donkey ", New CharacterFormat() With {
            .FontName = "Calibri",
            .Size = 16.0,
            .FontColor = Color.Black
        })

        ' Our picture has InlineLayout - it doesn't have positioning by coordinates
        ' and located as flowing content together with text (Run and other Inline elements).
        Dim pict1 As New Picture(dc, InlineLayout.Inline(New Size(100, 100)), pictPath)

        ' Add picture to the paragraph.
        par.Inlines.Add(pict1)

        ' Add some text content.
        par.Content.End.Insert(" arrive at Farquaad's palace in Duloc, where they end up in a tournament.", New CharacterFormat() With {
            .FontName = "Calibri",
            .Size = 16.0,
            .FontColor = Color.Black
        })

        ' 2. Picture with FloatingLayout:
        ' Floating layout means that the Picture (or Shape) is positioned by coordinates.
        Dim pict2 As New Picture(dc, pictPath)
        pict2.Layout = FloatingLayout.Floating(New HorizontalPosition(50, LengthUnit.Millimeter, HorizontalPositionAnchor.Page), New VerticalPosition(70, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), New Size(LengthUnitConverter.Convert(10, LengthUnit.Centimeter, LengthUnit.Point), LengthUnitConverter.Convert(10, LengthUnit.Centimeter, LengthUnit.Point)))

        ' Set the wrapping style.
        TryCast(pict2.Layout, FloatingLayout).WrappingStyle = WrappingStyle.Square

        ' Add our picture into the section.
        s.Content.End.Insert(pict2.Content)

        ' Save our document into DOCX format.
        dc.Save(documentPath, New DocxSaveOptions())

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub
End Module