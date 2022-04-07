Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            FormattingAndStyles()
        End Sub
        ''' <summary>
        ''' Creates a new document and applies formatting and styles.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/formatting-and-styles.php
        ''' </remarks>
        Private Shared Sub FormattingAndStyles()
            Dim docxPath As String = "FormattingAndStyles.docx"

            ' Let's create a new document.
            Dim dc As New DocumentCore()
            Dim run1 As New Run(dc, "This is Run 1 with character format Green. ")
            Dim run2 As New Run(dc, "This is Run 2 with style Red.")

            ' Create a new character style.
            Dim redStyle As New CharacterStyle("Red")
            redStyle.CharacterFormat.FontColor = Color.Red
            dc.Styles.Add(redStyle)

            ' Apply the direct character formatting.            
            run1.CharacterFormat.FontColor = Color.DarkGreen

            ' Apply only the style.
            run2.CharacterFormat.Style = redStyle

            dc.Content.End.Insert(run1.Content)
            dc.Content.End.Insert(run2.Content)

            ' Save our document into DOCX format.
            dc.Save(docxPath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(docxPath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
