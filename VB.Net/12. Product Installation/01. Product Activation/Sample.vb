Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        ProductActivation()
    End Sub

    ''' <summary>
    ''' Document .Net activation.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/product-activation.php
    ''' </remarks>
    Sub ProductActivation()
        ' Document .Net activation.

        ' You will get own serial number after purchasing the license.
        ' If you will have any questions, email us to sales@sautinsoft.com or ask at online chat https://www.sautinsoft.com.

        Dim serial As String = "1234567890"

        ' NOTICE: Place this line firstly, before creating of the DocumentCore object.
        DocumentCore.Serial = serial

        ' Let's create a new document by activated version.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert("Hello World!", New CharacterFormat() With {
                .FontName = "Verdana",
                .Size = 65.5F,
                .FontColor = Color.Orange
            })

        ' Save a document to a file in DOCX format.
        Dim filePath As String = "Result.docx"
        dc.Save(filePath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub
End Module