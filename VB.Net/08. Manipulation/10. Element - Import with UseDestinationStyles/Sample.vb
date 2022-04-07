Imports SautinSoft.Document
Imports SautinSoft.Document.Tables
Imports System.Linq

Namespace Sample
    Friend Class Sample

        Shared Sub Main(ByVal args() As String)
            ImportUseDestinationStyles()
        End Sub

        ''' <summary>
        ''' Import an Element with Styles from another document. Mode: UseDestinationStyles.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/import-element-use-destination-styles.php
        ''' </remarks>
        Private Shared Sub ImportUseDestinationStyles()
            ' Mode: UseDestinationStyles.

            ' 'UseDestinationStyles' means to copy only styles wchich are don't exist 
            ' in the destination document. 
            ' If the destination document already contains a style with the same name, 
            ' therefore a style from a source will not be copied.

            ' For example, a destination document contains a style "Green" (FontSize = 24, DarkGreen).
            ' And a source document also contains a style with name "Green" (FontSize = 20, Green, Underline).
            ' After the importing, the imported content will change its formatting correspondly to the "Green" style in the destination document.
            ' Because style "Green" (FontSize = 20, Green, Underline) will not be imported.

            Dim source As DocumentCore = DocumentCore.Load("..\..\..\SourceStyles.docx")
            Dim dest As New DocumentCore()

            ' Before importing a style from another document, let's create a style 
            ' with the same name but different formatting to see 
            ' how the name conflict will be resolved in mode 'UseDestinationStyles'.
            Dim chStyle As New CharacterStyle("Green")
            chStyle.CharacterFormat.FontColor = Color.DarkGreen
            chStyle.CharacterFormat.Size = 24
            dest.Styles.Add(chStyle)
            dest.Content.End.Insert((New Run(dest, "First ", New CharacterFormat() With {.Style = chStyle})).Content)

            ' Create an ImportSession with mode 'UseDestinationStyles'.
            Dim session As New ImportSession(source, dest, StyleImportingMode.UseDestinationStyles)

            ' Let's import a 1st paragraph from the source document.
            ' The paragraph contains a text marked by style "Green". 
            ' As a style with the same name is already exist, the new "Green" style will not be imported.
            Dim importedPar As Paragraph = dest.Import(Of Paragraph)(CType(source.Sections(0).Blocks(0), Paragraph), True, session)
            dest.Content.End.Insert(importedPar.Content)

            ' Save the destination document into DOCX format.
            Dim docPath As String = "UseDestinationStyles.docx"
            dest.Save(docPath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(docPath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
