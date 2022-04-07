Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports SautinSoft.Document.Tables
Imports System.Linq

Module Sample
    Sub Main()
        ElementManipulation()
    End Sub

    ''' <summary>
    ''' Create a document and add paragraphs as content and as element.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/element-manipulation.php
    ''' </remarks>
    Sub ElementManipulation()
        Dim filePath As String = "Result.docx"

        ' Let's create a new document.
        Dim dc As New DocumentCore()
        Dim par As New Paragraph(dc, "This is the first paragraph.")

        ' Insert the clone of our Paragraph using ContentRange.
        dc.Content.End.Insert(par.Content)

        ' Add our Paragraph in Block collection as Element.
        dc.Sections(0).Blocks.Add(par)

        ' Again, insert the clone of our Paragraph using ContentRange.
        dc.Content.End.Insert(par.Content)

        ' Change text in our Paragraph
        TryCast(par.Inlines(0), Run).Text = "Now we are in the second paragraph."

        ' Find 3rd paragraph and change text in it.
        TryCast((TryCast(par.NextSibling, Paragraph)).Inlines(0), Run).Text = "This is the third paragraph."

        ' Save our document.
        dc.Save(filePath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub
End Module