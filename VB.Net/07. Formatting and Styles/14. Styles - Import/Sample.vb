Imports SautinSoft.Document
Imports SautinSoft.Document.Tables
Imports System.Linq

Namespace Sample
    Friend Class Sample

        Shared Sub Main(ByVal args() As String)
            ImportStyles()
        End Sub

        ''' <summary>
        ''' This sample shows how to import styles from a one document to another. 
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/styles-import.php
        ''' </remarks>
        Public Shared Sub ImportStyles()
            Dim source As DocumentCore = DocumentCore.Load("..\..\..\document with styles.docx")
            Dim dest As New DocumentCore()
            dest.Sections.Add(New Section(dest))

            ' 1. Let's import all paragraph styles.
            For Each style As Style In source.Styles.Where(Function(s) TypeOf s Is ParagraphStyle)
                dest.Styles.Import(style)
            Next style

            ' 1.5. Let's import Character style with name from "Green".

            ' 1.5.1 Find a desired style with name "Green".
            Dim charStyle As CharacterStyle = CType(source.Styles.FirstOrDefault(Function(s) TypeOf s Is CharacterStyle AndAlso s.Name = "Green"), CharacterStyle)
            ' 1.5.2 Import the style into the "dest" document and get the reference to it.
            If charStyle IsNot Nothing Then
                charStyle = CType(dest.Styles.Import(charStyle), CharacterStyle)
            End If

            ' 2. Insert a new paragraph and apply the just now imported style.
            Dim p As New Paragraph(dest, "Charles Dickens was an extraordinary man. He is best known " & "as a novelist but he was very much more than that. He was as prominent in his other pursuits " & "but they were not areas of life where we can still see him today.")
            dest.Sections(0).Blocks.Add(p)

            ' Find a style with the name "Center"
            Dim pStyle As ParagraphStyle = CType(dest.Styles.FirstOrDefault(Function(s) TypeOf s Is ParagraphStyle AndAlso s.Name = "Center"), ParagraphStyle)

            ' Apply the style to the paragraph.
            If pStyle IsNot Nothing Then
                p.ParagraphFormat.Style = pStyle
            End If

            ' Aplly the character style "Green" to all Run element in our paragraph.
            If charStyle IsNot Nothing Then
                For Each r As Run In p.Inlines
                    r.CharacterFormat.Style = charStyle
                Next r
            End If

            ' Save the dest document into DOCX format.
            Dim docPath As String = "SimpleImport.docx"
            dest.Save(docPath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(docPath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
