Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        DeleteText()
    End Sub

    ''' <summary>
    ''' Delete a specific text from DOCX document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/delete-text-from-docx-document-net-csharp-vb.php
    ''' </remarks>
    Sub DeleteText()
        Dim filePath As String = "..\..\..\example.docx"
        Dim fileResult As String = "Result.pdf"
        Dim textToDelete As String = "document"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)

        Dim countDel As Integer = 0
        For Each cr As ContentRange In dc.Content.Find(textToDelete).Reverse()
            cr.Delete()
            countDel += 1
        Next cr
        Console.WriteLine("The text: """ & textToDelete & """ - was deleted " & countDel.ToString() & " time(s).")
        Console.ReadKey()

        dc.Save(fileResult)
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(fileResult) With {.UseShellExecute = True})
    End Sub
End Module