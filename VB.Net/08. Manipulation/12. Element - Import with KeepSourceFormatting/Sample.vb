Imports SautinSoft.Document
Imports SautinSoft.Document.Tables
Imports System.Linq

Namespace Sample
    Friend Class Sample

        Shared Sub Main(ByVal args() As String)
            ImportKeepSourceFormatting()
        End Sub

        ''' <summary>
        ''' Import an Element with Styles from another document. Mode: KeepSourceFormatting.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/import-element-keep-source-formatting.php
        ''' </remarks>
        Private Shared Sub ImportKeepSourceFormatting()
            ' Mode: KeepSourceFormatting.

            ' 'KeepSourceFormatting' means to copy all required styles to the destination document, 
            ' generate unique style names if needed.

            ' For example, a destination document contains a style "Green" (Calibri, FontSize = 20, Green, Underline).
            ' And a source document also contains an equal style with the same name "Green" (Calibri, FontSize = 20, Green, Underline).
            ' The style "Green" will be imported and renamed to "Green1".
            ' All imported elements linked to style "Green" will be remapped to style "Green1".

            Dim source As DocumentCore = DocumentCore.Load("..\..\..\SourceStyles.docx")
            Dim dest As New DocumentCore()

            ' Let's create a style "Green" (Calibri, FontSize = 20, Green, Underline).
            Dim chStyle As New CharacterStyle("Green")
            chStyle.CharacterFormat.FontName = "Calibri"
            chStyle.CharacterFormat.FontColor = Color.Green
            chStyle.CharacterFormat.Size = 20
            chStyle.CharacterFormat.UnderlineStyle = UnderlineType.Single
            dest.Styles.Add(chStyle)
            dest.Content.End.Insert((New Run(dest, "This text has the style Green.", New CharacterFormat() With {.Style = chStyle})).Content)

            ' Create an ImportSession with mode 'KeepSourceFormatting'.
            Dim session As New ImportSession(source, dest, StyleImportingMode.KeepSourceFormatting)

            ' Let's import a paragraph.
            ' The imported paragraph contains a text with style "Green" (FontSize = 20, Green, Underline). 
            ' The style "Green" will be imported and renamed to "Green1", because we already have "Green".
            ' All links in imported paragraph will be remapped to the style "Green1".
            Dim importedPar As Paragraph = dest.Import(Of Paragraph)(CType(source.Sections(0).Blocks(0), Paragraph), True, session)
            dest.Content.End.Insert(importedPar.Content)

            ' Save the destination document into DOCX format.
            Dim docPath As String = "KeepSourceFormatting.docx"
            dest.Save(docPath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(docPath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace