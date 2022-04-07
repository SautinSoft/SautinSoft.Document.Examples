Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        ReplaceHeader()
    End Sub

    ''' <summary>
    ''' Removes the old header/footer and inserts a new one into an existing PDF document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/remove-header-and-footer-in-pdf-net-csharp-vb.php
    ''' </remarks>
    Sub ReplaceHeader()
        Dim inpFile As String = "..\..\..\somebody.pdf"
        Dim outFile As String = "With new Header.pdf"
        Dim dc As DocumentCore = DocumentCore.Load(inpFile)

        ' Create new header with formatted text.
        Dim header As New HeaderFooter(dc, HeaderFooterType.HeaderDefault)
        header.Content.Start.Insert("Modified : 1 April 2020", New CharacterFormat() With {
                .Size = 14.0,
                .FontColor = Color.DarkGreen
            })
        ' Add 10 mm from Top before new header.
        TryCast(header.Blocks(0), Paragraph).ParagraphFormat.SpaceBefore = LengthUnitConverter.Convert(10, LengthUnit.Millimeter, LengthUnit.Point)


        For Each s As Section In dc.Sections
            ' Find the first paragraph (Let's assume that it's the header) and remove it.
            If s.Blocks.Count > 0 Then
                s.Blocks.RemoveAt(0)
            End If
            ' Insert the new header into the each section.
            s.HeadersFooters.Add(header.Clone(True))
        Next s

        dc.Save(outFile)

        ' Open the results for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(inpFile) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
    End Sub
End Module