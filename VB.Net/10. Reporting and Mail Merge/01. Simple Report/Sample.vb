Option Infer On

Imports System
Imports System.IO
Imports SautinSoft.Document
Module Sample
    Sub Main()
        MailMergeSimpleEnvelope()
    End Sub

    ''' <summary>
    ''' Generates 5 envelopes "Happy New Year" for Simpson family using the one template.
    ''' </summary>
    ''' <remarks>
    ''' See details at: https://sautinsoft.com/products/document/help/net/developer-guide/mail-merge-simple-report-net-csharp-vb.php
    ''' </remarks>
    Sub MailMergeSimpleEnvelope()
        Dim templatePath As String = "..\..\..\envelope-template.docx"
        Dim resultPath As String = "Simpson-family.docx"

        Dim dc As DocumentCore = DocumentCore.Load(templatePath)

        Dim dataSource = {
            New With {
                Key .Name = "Homer",
                Key .FamilyName = "Simpson"
            },
            New With {
                Key .Name = "Marge ",
                Key .FamilyName = "Simpson"
            },
            New With {
                Key .Name = "Bart",
                Key .FamilyName = "Simpson"
            },
            New With {
                Key .Name = "Lisa",
                Key .FamilyName = "Simpson"
            },
            New With {
                Key .Name = "Maggie",
                Key .FamilyName = "Simpson"
            }
        }

        dc.MailMerge.Execute(dataSource)
        dc.Save(resultPath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
    End Sub

End Module