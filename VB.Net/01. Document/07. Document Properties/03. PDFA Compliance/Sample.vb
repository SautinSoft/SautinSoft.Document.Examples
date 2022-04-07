Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        LoadAndSaveAsPDFA()
    End Sub

    ''' <summary>
    ''' Load an existing document (*.docx, *.rtf, *.pdf, *.html, *.txt, *.pdf) and save it as a PDF/A compliant version. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/create-and-save-document-in-pdf-a-format-net-csharp-vb.php
    ''' </remarks>
    Sub LoadAndSaveAsPDFA()
        ' Path to a loadable document.
        Dim loadPath As String = "..\..\..\example.docx"

        Dim dc As DocumentCore = DocumentCore.Load(loadPath)

        Dim options As New PdfSaveOptions() With {.Compliance = PdfCompliance.PDF_A1a}

        Dim savePath As String = Path.ChangeExtension(loadPath, ".pdf")
        dc.Save(savePath, options)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(savePath) With {.UseShellExecute = True})
    End Sub
End Module