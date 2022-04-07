Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Sample
    Friend Class Sample
        Shared Sub Main(ByVal args() As String)
            FindAndReplace()
        End Sub
        ''' <summary>
        ''' Find and replace a specific text in an existing DOCX document.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/find-and-replace-text-in-docx-document-net-csharp-vb.php
        ''' </remarks>
        Public Shared Sub FindAndReplace()
            ' Path to the loadable document.
            Dim loadPath As String = "..\..\..\example.docx"

            Dim dc As DocumentCore = DocumentCore.Load(loadPath)

            ' Find "Bean" and Replace everywhere on "Joker"
            Dim regex As New Regex("Bean", RegexOptions.IgnoreCase)

            ' Start:

            ' Please note, Reverse() makes sure that action Replace() doesn't affect to Find().
            For Each item As ContentRange In dc.Content.Find(regex).Reverse()
                item.Replace("Joker", New CharacterFormat() With {
                .BackgroundColor = Color.Yellow,
                .FontName = "Arial",
                .Size = 16.0
                   })
            Next item

            ' End:

            ' The code above finds and replaces the content in the whole document.
            ' Let us say, you want to replace a text inside shape blocks only:

            ' 1. Comment the code above from the line "Start" to the "End".
            ' 2. Uncomment this code:
            'For Each shp As Shape In dc.GetChildElements(True, ElementType.Shape).Reverse()
            'For Each item As ContentRange In shp.Content.Find(regex).Reverse()
            'item.Replace("Joker", New CharacterFormat() With {
            '.BackgroundColor = Color.Yellow,
            '.FontName = "Arial",
            '.Size = 16.0
            '       })
            'Next item
            'Next shp

            ' Save the document as DOCX format.
            Dim savePath As String = Path.ChangeExtension(loadPath, ".replaced.docx")
            dc.Save(savePath, SaveOptions.DocxDefault)

            ' Open the original and result documents for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(loadPath) With {.UseShellExecute = True})
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(savePath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
