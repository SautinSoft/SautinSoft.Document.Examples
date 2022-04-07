Imports SautinSoft.Document
Imports SautinSoft.Document.Tables
Imports System.Linq

Namespace Sample
    Friend Class Sample

        Shared Sub Main(ByVal args() As String)
            ImportKeepDifferentStyles()
        End Sub

        ''' <summary>
        ''' Import an Element with Styles from another document. Mode: KeepDifferentStyles.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/import-element-keep-different-styles.php
        ''' </remarks>
        Private Shared Sub ImportKeepDifferentStyles()
            ' Mode: KeepDifferentStyles 
            ' The most useful mode to preserve formatting during importing.

            ' 'KeepDifferentStyles' means to only copy styles that are different by formatting.
            ' If the destination document already contains a style with the same name, 
            ' therefore an unique style name will be generated.

            ' For example, a destination document contains a style "MyGreen" (FontSize = 20, Green, Underline).
            ' And a source document also contains a style with name "Green" (FontSize = 20, Green, Underline).
            ' Because of the formatting of styles are equal, the "Green" style will not be imported.
            ' All imported elements linked to style "Green" will be remapped to style "MyGreen"

            Dim source As DocumentCore = DocumentCore.Load("..\..\..\SourceStyles.docx")
            Dim dest As New DocumentCore()

            ' Create a new style "MyGreen" (FontSize = 20, Green, Underline).
            Dim chStyle As New CharacterStyle("MyGreen")
            chStyle.CharacterFormat.FontColor = Color.Green
            chStyle.CharacterFormat.Size = 20
            chStyle.CharacterFormat.FontName = "Calibri"
            chStyle.CharacterFormat.UnderlineStyle = UnderlineType.Single
            dest.Styles.Add(chStyle)
            dest.Content.End.Insert((New Run(dest, "This text has the style MyGreen.", New CharacterFormat() With {.Style = chStyle})).Content)

            ' Create an ImportSession with mode 'KeepDifferentStyles'.
            Dim session As New ImportSession(source, dest, StyleImportingMode.KeepDifferentStyles)

            ' Let's import a paragraph.
            ' The imported paragraph contains a text with style "Green" (FontSize = 20, Green, Underline). 
            ' The style "Green" will not be imported, because we already have "MyGreen" with the same formatting.
            ' All links to "Green" will be remapped to "MyGreen".
            Dim importedPar As Paragraph = dest.Import(Of Paragraph)(CType(source.Sections(0).Blocks(0), Paragraph), True, session)
            dest.Content.End.Insert(importedPar.Content)

            ' Save the destination document into DOCX format.
            Dim docPath As String = "KeepDifferentStyles.docx"
            dest.Save(docPath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(docPath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
