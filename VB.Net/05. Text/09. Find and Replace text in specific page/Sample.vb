Option Infer On
Imports System
Imports System.IO
Imports SautinSoft.Document
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Sample
    Friend Class Sample
        Shared Sub Main(ByVal args() As String)
            FindAndReplaceSpecificTextWay1()
            'FindAndReplaceSpecificTextWay2();
        End Sub
        ''' <summary>
        ''' Find and replace a certain text only on the second page of the DOCX document.
        ''' Way #1:
        ''' The DOCX document is loaded, then the analysis of the number of pages inside (GetPaginator).
        ''' DOCX is not a "paged" format and has to be paginated.
        ''' There may be problems with the text in the tables, since the transfer to a new page is possible.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-text-in-specific-page-after-specific-text-net-csharp-vb.php
        ''' </remarks>
        Public Shared Sub FindAndReplaceSpecificTextWay1()
            Dim inpFileWord As String = "..\..\..\example.docx"
            Dim outFileWord As String = "result1.docx"
            Dim searchText = "italic"

            Dim dc As DocumentCore = DocumentCore.Load(inpFileWord)
            Dim regex As New Regex(searchText, RegexOptions.IgnoreCase)

            Dim dp As DocumentPaginator = dc.GetPaginator()

            If dp.Pages.Count > 2 Then
                ' Find and replace a certain text only on the second page of the DOCX document.
                ' If you need the first page - 0, the third page - 2, etc.
                For Each item As ContentRange In dp.Pages(1).Content.Find(regex).Reverse()
                    item.Replace("This word has been corrected", New CharacterFormat() With {
                        .BackgroundColor = Color.Yellow,
                        .FontName = "Arial",
                        .Size = 16.0
                    })
                Next item
            End If
            dc.Save(outFileWord, New DocxSaveOptions())
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(inpFileWord) With {.UseShellExecute = True})
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFileWord) With {.UseShellExecute = True})
        End Sub

        ''' <summary>
        ''' Find and replace a certain text only on the second page of the DOCX document.
        ''' Way #2:
        ''' The DOCX document is loaded. DOCX is not a "paged" format and has to be paginated.
        ''' We will convert DOCX to PDF and then to save in DOCX back.
        ''' There may be problems with Formatting, because we are doing reverse format conversion.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-text-in-specific-page-after-specific-text-net-csharp-vb.php
        ''' </remarks>
        Public Shared Sub FindAndReplaceSpecificTextWay2()
            Dim inpFileWord As String = "..\..\..\example.docx"
            Dim tempPDFFile As String = "example_temp.pdf"
            Dim outFileWord As String = "result2.docx"
            Dim searchText = "italic"

            Dim dc1 As DocumentCore = DocumentCore.Load(inpFileWord)
            dc1.Save(tempPDFFile)
            Dim dc2 As DocumentCore = DocumentCore.Load(tempPDFFile)
            Dim regex As New Regex(searchText, RegexOptions.IgnoreCase)

            If True Then
                For index As Integer = 0 To dc2.Sections.Count - 1
                    Dim page = dc2.Sections(index)
                    If dc2.Sections.Count > 2 Then
                        ' Find and replace a certain text only on the second page of the DOCX document.
                        ' If you need the first page - 0, the third page - 2, etc.
                        For Each item As ContentRange In dc2.Sections(1).Content.Find(regex).Reverse()
                            item.Replace("This word has been corrected", New CharacterFormat() With {
                                .BackgroundColor = Color.Yellow,
                                .FontName = "Arial",
                                .Size = 16.0
                            })
                        Next item
                    End If
                Next index
                dc2.Save(outFileWord)
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(inpFileWord) With {.UseShellExecute = True})
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFileWord) With {.UseShellExecute = True})
            End If
        End Sub
    End Class
End Namespace