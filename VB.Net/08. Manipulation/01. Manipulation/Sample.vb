Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        Manipulation()
    End Sub

    ''' <summary>
    ''' Replace all Run elements with Bold formatting to Italic and mark them by yellow.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/manipulation.php
    ''' </remarks>
    Sub Manipulation()
        Dim filePath As String = "..\..\..\example.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        Dim filePathResult As String = "Result-file.pdf"

        For Each run As Run In dc.GetChildElements(True, ElementType.Run)
            If run.CharacterFormat.Bold = True Then
                run.CharacterFormat.Bold = False
                run.CharacterFormat.Italic = True
                run.CharacterFormat.BackgroundColor = Color.Yellow
            End If
        Next run
        dc.Save(filePathResult)
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePathResult) With {.UseShellExecute = True})
    End Sub
End Module