Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        SaveToHtmlFile()
        SaveToHtmlStream()
    End Sub

    ''' <summary>
    ''' Open an existing document and saves it as HTML files (in the Fixed and Flowing modes).
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-html-net-csharp-vb.php
    ''' </remarks>
    Sub SaveToHtmlFile()
        Dim inputFile As String = "..\..\..\example.docx"

        Dim dc As DocumentCore = DocumentCore.Load(inputFile)

        Dim fileHtmlFixed As String = "Fixed-as-file.html"
        Dim fileHtmlFlowing As String = "Flowing-as-file.html"

        ' Save to HTML file: HtmlFixed.
        dc.Save(fileHtmlFixed, New HtmlFixedSaveOptions() With {
            .Version = HtmlVersion.Html5,
            .CssExportMode = CssExportMode.Inline
        })

        ' Save to HTML file: HtmlFlowing.
        dc.Save(fileHtmlFlowing, New HtmlFlowingSaveOptions() With {
            .Version = HtmlVersion.Html5,
            .CssExportMode = CssExportMode.Inline,
            .ListExportMode = HtmlListExportMode.ByHtmlTags
        })

        ' Open the results for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(fileHtmlFixed) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(fileHtmlFlowing) With {.UseShellExecute = True})

    End Sub

    ''' <summary>
    ''' Creates a new document and saves it as HTML documents (in the Fixed and Flowing modes) using MemoryStream.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-html-net-csharp-vb.php
    ''' </remarks>
    Sub SaveToHtmlStream()
        ' There variables are necessary only for demonstration purposes.
        Dim fileData() As Byte = Nothing
        Dim fileHtmlFixed As String = "Fixed-as-stream.html"
        Dim fileHtmlFlowing As String = "Flowing-as-stream.html"

        ' Assume we already have a document 'dc'.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert("Hey Guys and Girls!")

        ' Let's save our document to a MemoryStream.
        Using ms As New MemoryStream()
            ' HTML Fixed.
            dc.Save(ms, New HtmlFixedSaveOptions())
            fileData = ms.ToArray()

            File.WriteAllBytes(fileHtmlFixed, fileData)

            ' Or HTML flowing.
            dc.Save(ms, New HtmlFlowingSaveOptions())
            fileData = ms.ToArray()

            File.WriteAllBytes(fileHtmlFlowing, fileData)
        End Using
    End Sub
End Module