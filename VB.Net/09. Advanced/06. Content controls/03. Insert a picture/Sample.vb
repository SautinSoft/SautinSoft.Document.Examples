Option Infer On
Imports System.Threading.Tasks
Imports SautinSoft.Document
Imports SautinSoft.Document.CustomMarkups
Imports SautinSoft.Document.Drawing
Imports SautinSoft.Document.Tables
Imports System.Net
Imports System.Net.Http
Imports System.IO



Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            InsertPicture()
        End Sub
        ''' <summary>
        ''' Inserting a Picture content control.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/content-controls-insert-picture-net-csharp-vb.php
        ''' </remarks>

        Private Shared Sub InsertPicture()
            ' Let's create a simple document.
            Dim dc As New DocumentCore()
            Dim pict, pict1, pict2 As Picture

            Dim imageBytes() As Byte = File.ReadAllBytes("..\..\..\banner_sautinsoft.jpg")
            Using ms As New MemoryStream(imageBytes)
                pict = New Picture(dc, New InlineLayout(New Size(400, 100)), ms)
            End Using

            imageBytes = File.ReadAllBytes("..\..\..\developer.png")
            Using ms As New MemoryStream(imageBytes)
                pict1 = New Picture(dc, New InlineLayout(New Size(50, 50)), ms)
            End Using

            Dim section As New Section(dc)
            dc.Sections.Add(section)

            section.Blocks.Add(New Paragraph(dc, "Picture below is inside the block-level content control:"))

            ' Create a picture content control.
            Dim control As New BlockContentControl(dc, ContentControlType.Picture, New Paragraph(dc, pict))
            section.Blocks.Add(control)

            Dim par As New Paragraph(dc, New Run(dc, "Following picture is inside the inline-level content control: "), New InlineContentControl(dc, ContentControlType.Picture, pict1))
            section.Blocks.Add(par)

            section.Blocks.Add(New Paragraph(dc, "Insert a picture content control from the local disk:"))
            Dim pictPath As String = "..\..\..\picture.jpg"
            pict2 = New Picture(dc, New InlineLayout(New Size(100, 100)), pictPath)
            Dim localpict As New BlockContentControl(dc, ContentControlType.Picture, New Paragraph(dc, pict2))
            section.Blocks.Add(localpict)

            ' Save our document into DOCX format.
            Dim resultPath As String = "result.docx"
            dc.Save(resultPath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace