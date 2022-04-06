using System.IO;
using System.Linq;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            DigitalSignature();
        }

        /// <summary>
        /// Load an existing document (*.docx, *.rtf, *.pdf, *.html, *.txt, *.pdf) and save it in a PDF document with the digital signature.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/digital-signature-net-csharp-vb.php
        /// </remarks>
        public static void DigitalSignature()
        {
            // Path to a loadable document.
            string loadPath = @"..\..\..\digitalsignature.docx";
            string savePath = "Result.pdf";

            DocumentCore dc = DocumentCore.Load(loadPath);

            // Create a new invisible Shape for the digital signature.       
            // Place the Shape into top-left corner (0 mm, 0 mm) of page.
            Shape signatureShape = new Shape(dc, Layout.Floating(new HorizontalPosition(0f, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin),
                                    new VerticalPosition(0f, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), new Size(1, 1)));
            ((FloatingLayout)signatureShape.Layout).WrappingStyle = WrappingStyle.InFrontOfText;
            signatureShape.Outline.Fill.SetEmpty();

            // Find a first paragraph and insert our Shape inside it.
            Paragraph firstPar = dc.GetChildElements(true).OfType<Paragraph>().FirstOrDefault();
            firstPar.Inlines.Add(signatureShape);

            // Picture which symbolizes a handwritten signature.
            Picture signaturePict = new Picture(dc, @"..\..\..\signature.png");

            // Signature picture will be positioned:
            // 14.5 cm from Top of the Shape.
            // 4.5 cm from Left of the Shape.
            signaturePict.Layout = Layout.Floating(
               new HorizontalPosition(4.5, LengthUnit.Centimeter, HorizontalPositionAnchor.Page),
               new VerticalPosition(14.5, LengthUnit.Centimeter, VerticalPositionAnchor.Page),
               new Size(20, 10, LengthUnit.Millimeter));

            PdfSaveOptions options = new PdfSaveOptions();

            // Path to the certificate (*.pfx).
            options.DigitalSignature.CertificatePath = @"..\..\..\sautinsoft.pfx";

            // The password for the certificate.
            // Each certificate is protected by a password.
            // The reason is to prevent unauthorized the using of the certificate.
            options.DigitalSignature.CertificatePassword = "123456789";

            // Additional information about the certificate.
            options.DigitalSignature.Location = "World Wide Web";
            options.DigitalSignature.Reason = "Document.Net by SautinSoft";
            options.DigitalSignature.ContactInfo = "info@sautinsoft.com";

            // Placeholder where signature should be visualized.
            options.DigitalSignature.SignatureLine = signatureShape;

            // Visual representation of digital signature.
            options.DigitalSignature.Signature = signaturePict;

            dc.Save(savePath, options);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(savePath) { UseShellExecute = true });
        }
    }
}