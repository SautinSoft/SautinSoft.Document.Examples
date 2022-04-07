Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        ChangeHeaderAndFooter()
    End Sub

    ''' <summary>
    ''' How to edit Header and Footer in PDF file
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/edit-header-and-footer-in-pdf-net-csharp-vb.php
    ''' </remarks>
    Sub ChangeHeaderAndFooter()
        Dim inpFile As String = "..\..\..\somebody.pdf"
        Dim outFile As String = "With changed header and footer.pdf"
        Dim dc As DocumentCore = DocumentCore.Load(inpFile)

        ' Create new header with formatted text.
        Dim header As New HeaderFooter(dc, HeaderFooterType.HeaderDefault)
        header.Content.Start.Insert("Modified : 1 April 2020", New CharacterFormat() With {
                .Size = 14.0,
                .FontColor = Color.DarkGreen
            })

        ' Create the footer with orange text, with font name Elephant and size of 50 pt.
        Dim footer As New HeaderFooter(dc, HeaderFooterType.FooterDefault)
        Dim p As Paragraph = New Paragraph(dc, New Run(dc, "Last modified: 1st June 2021", New CharacterFormat() With {
                .Size = 50.0,
                .FontColor = Color.Orange,
                .FontName = "Elephant"
            }))
        p.ParagraphFormat.Alignment = HorizontalAlignment.Left
        footer.Blocks.Add(p)

        For Each s As Section In dc.Sections
            If s.Blocks.Count > 0 Then
                s.Blocks.RemoveAt(1)
            End If
            s.HeadersFooters.Add(header.Clone(True))
            s.HeadersFooters.Add(footer.Clone(True))
        Next s
        dc.Save(outFile)

        ' Open the results for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(inpFile) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
    End Sub
End Module