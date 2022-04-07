Imports System
Imports System.Linq
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Module Sample
    Sub Main()
        DigitalSignature()
    End Sub

    ''' <summary>
    ''' Load an existing document (*.docx, *.rtf, *.pdf, *.html, *.txt, *.pdf) and save it in a PDF document with the digital signature.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/digital-signature-net-csharp-vb.php
    ''' </remarks>
    Sub DigitalSignature()
        ' Path to a loadable document.
        Dim loadPath As String = "..\..\..\digitalsignature.docx"
        Dim savePath As String = "Result.pdf"

        Dim dc As DocumentCore = DocumentCore.Load(loadPath)

        ' Create a new invisible Shape for the digital signature.       
        ' Place the Shape into top-left corner (0 mm, 0 mm) of page.
        Dim signatureShape As New Shape(dc, Layout.Floating(New HorizontalPosition(0F, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin), New VerticalPosition(0F, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), New Size(1, 1)))
        CType(signatureShape.Layout, FloatingLayout).WrappingStyle = WrappingStyle.InFrontOfText
        signatureShape.Outline.Fill.SetEmpty()

        ' Find a first paragraph and insert our Shape inside it.
        Dim firstPar As Paragraph = dc.GetChildElements(True).OfType(Of Paragraph)().FirstOrDefault()
        firstPar.Inlines.Add(signatureShape)

        ' Picture which symbolizes a handwritten signature.
        Dim signaturePict As New Picture(dc, "..\..\..\signature.png")

        ' Signature picture will be positioned:
        ' 14.5 cm from Top of the Shape.
        ' 4.5 cm from Left of the Shape.
        signaturePict.Layout = Layout.Floating(New HorizontalPosition(4.5, LengthUnit.Centimeter, HorizontalPositionAnchor.Page), New VerticalPosition(14.5, LengthUnit.Centimeter, VerticalPositionAnchor.Page), New Size(20, 10, LengthUnit.Millimeter))

        Dim options As New PdfSaveOptions()

        ' Path to the certificate (*.pfx).
        options.DigitalSignature.CertificatePath = "..\..\..\sautinsoft.pfx"

        ' The password for the certificate.
        ' Each certificate is protected by a password.
        ' The reason is to prevent unauthorized the using of the certificate.
        options.DigitalSignature.CertificatePassword = "123456789"

        ' Additional information about the certificate.
        options.DigitalSignature.Location = "World Wide Web"
        options.DigitalSignature.Reason = "Document.Net by SautinSoft"
        options.DigitalSignature.ContactInfo = "info@sautinsoft.com"

        ' Placeholder where signature should be visualized.
        options.DigitalSignature.SignatureLine = signatureShape

        ' Visual representation of digital signature.
        options.DigitalSignature.Signature = signaturePict

        dc.Save(savePath, options)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(savePath) With {.UseShellExecute = True})
    End Sub
End Module