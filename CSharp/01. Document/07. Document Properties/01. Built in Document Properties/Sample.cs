using System;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.IO;
using System.Linq;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateDocumentProperties();
            ReadDocumentProperties();
        }
		
        /// <summary>
        /// Create a new document (DOCX) with some built-in properties.
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/document-properties.php
        /// </remarks>
        public static void CreateDocumentProperties()
        {
            string filePath = @"..\..\..\DocumentProperties.docx";

            DocumentCore dc = new DocumentCore();

            // Let's create a simple inscription.
            dc.Content.End.Insert("Hello World!!!", new CharacterFormat() { FontName = "Verdana", Size = 65.5f, FontColor = Color.Orange });

            // Let's add some documents properties: Author, Subject, Company.
            dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Title] = "How to add document properties. It works with DOCX, RTF, PDF, HTML etc";
            dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Company] = "SautinSoft";
            dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Author] = "John Smith";
            dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Subject] = "Document .Net";
            dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Keywords] = "reader, writer, docx, pdf, html, rtf, text";
            dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.HyperlinkBase] = "www.sautinsoft.com";
            dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Manager] = "Alex Dickard";
            dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Category] = "Document Object Model (DOM)";
            dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.DateContentCreated] =
                new DateTime(2010, 1, 10).ToString();
            dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.DateLastSaved] =
                DateTime.Now.ToString();

            dc.CalculateStats();
            
            // Save our document to DOCX format.
            dc.Save(filePath);
			
            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }

        /// <summary>
        /// Read built-in document properties (from .docx) and enumerate them in new PDF document as small report.
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/document-properties.php
        /// </remarks>
        public static void ReadDocumentProperties()
        {
            string inpFile = @"..\..\..\DocumentProperties.docx";
            string statFile = @"..\..\..\Statistics.pdf";

            DocumentCore dc = DocumentCore.Load(inpFile);
			
            // Let's add some additional information. It can be anything you like.
            dc.Document.Properties.Custom.Add("Producer", "My Producer");

            // Add a paragraph in which all standard information about the document will be stored.
            Paragraph builtInPara = new Paragraph(dc,
            new Run(dc, "Built-in document properties:"),
            new SpecialCharacter(dc, SpecialCharacterType.LineBreak));
            builtInPara.ParagraphFormat.Alignment = HorizontalAlignment.Left;

            foreach (var docProp in dc.Document.Properties.BuiltIn)
            {
                builtInPara.Inlines.Add(
                    new Run(dc, string.Format("{0}: {1}", docProp.Key, docProp.Value)));

                builtInPara.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.LineBreak));
            }
			
            // Add a paragraph in which all additional information about the document will be stored.
            Paragraph customPropPara = new Paragraph(dc,
               new Run(dc, "Custom document properties:"),
               new SpecialCharacter(dc, SpecialCharacterType.LineBreak));
            customPropPara.ParagraphFormat.Alignment = HorizontalAlignment.Left;

            foreach (var docProp in dc.Document.Properties.Custom)
            {
                customPropPara.Inlines.Add(
                    new Run(dc, string.Format("{0}: {1} (Type: {2})", docProp.Key, docProp.Value, docProp.Value.GetType())));

                customPropPara.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.LineBreak));
            }
			
            // Add all document properties in the document and save it as PDF file.
            dc.Sections.Clear();
            dc.Sections.Add(new Section(dc, builtInPara, customPropPara));

            dc.Save(statFile, new PdfSaveOptions());
			
            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(statFile) { UseShellExecute = true });
        }
    }
}