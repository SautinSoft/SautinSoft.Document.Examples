Imports SautinSoft
Imports SautinSoft.Document
Imports System
Imports System.IO
Imports System.Reflection

Namespace Sample
    Class Sample
        ''' <summary>
        ''' Convert Word to PDF (Factur-X) using VB.NET and .NET.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-word-to-pdfa-factur-x.php
        ''' </remarks>
        Shared Sub Main(args As String())
            Dim inpFile As String = "..\..\..\example.docx"
            Dim xmlInfo As String = File.ReadAllText("..\..\..\Factur-X\Factur.xml")

            Dim outFile As String = "..\..\..\FacturXResult.pdf"

            Dim dc As DocumentCore = DocumentCore.Load(inpFile)

            Dim pdfSO As New PdfSaveOptions() With {
                .FacturXXML = xmlInfo
            }

            dc.Save(outFile, pdfSO)

            ' Important for Linux: Install MS Fonts
            ' sudo apt install ttf-mscorefonts-installer -y

            ' Open the result for demonstration purposes.
            Dim psi As New System.Diagnostics.ProcessStartInfo(outFile) With {
                .UseShellExecute = True
            }
            System.Diagnostics.Process.Start(psi)
        End Sub
    End Class
End Namespace