Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        SecureDocument()
    End Sub

    ''' <summary>
    ''' Create and secure a PDF document by password. Also set the permissions for the document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/security-options-net-csharp-vb.php
    ''' </remarks>
    Sub SecureDocument()
        Dim filePath As String = "ProtectedDocument.pdf"

        Dim dc As New DocumentCore()

        ' Let's create a simple document.
        dc.Content.End.Insert("Hello World!!!", New CharacterFormat() With {
                .FontName = "Verdana",
                .Size = 65.5F,
                .FontColor = Color.Orange
            })

        Dim so As New PdfSaveOptions()
        ' Password Protection
        so.EncryptionDetails.UserPassword = "12345"
        ' EncryptionAlgorithm
        so.EncryptionDetails.EncryptionAlgorithm = PdfEncryptionAlgorithm.RC4_128
        'Permissions: Content Copying, Commenting, Printing, Changing the Document, filing of form fildes etc
        'Printing: Allowed
        so.EncryptionDetails.Permissions = PdfPermissions.Printing

        ' Save a document as the PDF file with Security Options.
        dc.Save(filePath, so)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub
End Module