Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        AddHeaderFooter()
    End Sub

    ''' <summary>
    ''' How to add a header and footer into PDF document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/add-header-and-footer-in-pdf-net-csharp-vb.php
    ''' </remarks>
    Sub AddHeaderFooter()
        Dim inpFile As String = "..\..\..\shrek.pdf"
        Dim outFile As String = "Shrek with header and footer.pdf"
        Dim dc As DocumentCore = DocumentCore.Load(inpFile)

        ' Create new header with formatted text.
        Dim header As New HeaderFooter(dc, HeaderFooterType.HeaderDefault)
        header.Content.Start.Insert("Shrek and Donkey", New CharacterFormat() With {
                .Size = 14.0,
                .FontColor = Color.Brown
            })
        For Each s As Section In dc.Sections
            s.HeadersFooters.Add(header.Clone(True))
        Next s

        ' Create new footer with formatted text.
        Dim footer As New HeaderFooter(dc, HeaderFooterType.FooterDefault)
        footer.Content.Start.Insert("Fiona.", New CharacterFormat() With {
                .Size = 14.0,
                .FontColor = Color.Blue
            })
        For Each s As Section In dc.Sections
            s.HeadersFooters.Add(footer.Clone(True))
        Next s

        dc.Save(outFile)

        ' Open the PDF documents for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(inpFile) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
    End Sub
End Module