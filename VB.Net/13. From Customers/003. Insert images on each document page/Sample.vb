Imports System
Imports System.Collections.Generic
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Namespace Sample
    Friend Class Sample
        Shared Sub Main(ByVal args() As String)
            InsertImagesOnEachPage()
        End Sub
        ''' <summary>
        ''' Insert images on each page of the PDF file.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-insert-images-on-each-document-page-csharp-vb-net.php
        ''' </remarks>
        Public Shared Sub InsertImagesOnEachPage()
            Dim inpfFile As String = "..\..\..\example.pdf"
            Dim pctFile As String = "..\..\..\signature.png"
            Dim outFile As String = "Result.pdf"

            ' This example is acceptable for PDF documents.
            ' Because when we're loading PDF documents using DocumentCore the each PDF-page
            ' we'll be stored in own Section object.
            ' In other words, the each Section represents the separate PDF-page.
            Dim dc As DocumentCore = DocumentCore.Load(inpfFile)

            ' Load the Picture from a file.
            Dim pict As New Picture(dc, pctFile)

            ' In this example we'll place the image in three (3) 
            ' different places to the each document page.

            ' Let's create three layouts
            Dim layouts As New List(Of Layout)() From {FloatingLayout.Floating(New HorizontalPosition(10, LengthUnit.Millimeter, HorizontalPositionAnchor.Page), New VerticalPosition(260, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), New Size(LengthUnitConverter.Convert(2, LengthUnit.Centimeter, LengthUnit.Point), LengthUnitConverter.Convert(1, LengthUnit.Centimeter, LengthUnit.Point))), FloatingLayout.Floating(New HorizontalPosition(180, LengthUnit.Millimeter, HorizontalPositionAnchor.Page), New VerticalPosition(10, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), New Size(LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.Point), LengthUnitConverter.Convert(2, LengthUnit.Centimeter, LengthUnit.Point))), FloatingLayout.Floating(New HorizontalPosition(150, LengthUnit.Millimeter, HorizontalPositionAnchor.Page), New VerticalPosition(150, LengthUnit.Millimeter, VerticalPositionAnchor.Page), New Size(LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.Point), LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.Point)))}

            ' Iterate by Sections (PDF pages in our case).
            For Each s As Section In dc.Sections
                ' Insert our pictures in different places.
                For Each fl As FloatingLayout In layouts
                    pict.Layout = New FloatingLayout(fl.HorizontalPosition, fl.VerticalPosition, fl.Size)
                    ' Place the picture behind the text.
                    TryCast(pict.Layout, FloatingLayout).WrappingStyle = WrappingStyle.BehindText

                    ' Here we insert the Picture content at the 1st Block element (Paragraph or Table).
                    s.Blocks(0).Content.Start.Insert(pict.Content)
                Next fl
            Next s

            dc.Save(outFile, New PdfSaveOptions())

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
