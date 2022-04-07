Option Infer On

Imports System
Imports System.IO
Imports SautinSoft.Document
Imports System.Linq

Module Sample
    Sub Main()
        CreateDocumentProperties()
        ReadDocumentProperties()
    End Sub

    ''' <summary>
    ''' Create a new document (DOCX) with some built-in properties.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/document-properties.php
    ''' </remarks>
    Sub CreateDocumentProperties()
        Dim filePath As String = "..\..\..\DocumentProperties.docx"

        Dim dc As New DocumentCore()

        ' Let's create a simple inscription.
        dc.Content.End.Insert("Hello World!!!", New CharacterFormat() With {
            .FontName = "Verdana",
            .Size = 65.5F,
            .FontColor = Color.Orange
        })

        ' Let's add some documents properties: Author, Subject, Company.
        dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.Title) = "How to add document properties. It works with DOCX, RTF, PDF, HTML etc"
        dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.Company) = "SautinSoft"
        dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.Author) = "John Smith"
        dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.Subject) = "Document .Net"
        dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.Keywords) = "reader, writer, docx, pdf, html, rtf, text"
        dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.HyperlinkBase) = "www.sautinsoft.com"
        dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.Manager) = "Alex Dickard"
        dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.Category) = "Document Object Model (DOM)"
        dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.DateContentCreated) = (New Date(2010, 1, 10)).ToString()
        dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.DateLastSaved) = Date.Now.ToString()

        dc.CalculateStats()

        ' Save our document to DOCX format.
        dc.Save(filePath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub

    ''' <summary>
    ''' Read built-in document properties (from .docx) and enumerate them in new PDF document as small report.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/document-properties.php
    ''' </remarks>
    Sub ReadDocumentProperties()
        Dim inpFile As String = "..\..\..\DocumentProperties.docx"
        Dim statFile As String = "..\..\..\Statistics.pdf"

        Dim dc As DocumentCore = DocumentCore.Load(inpFile)

        ' Let's add some additional inforamtion. It can be anything you like.
        dc.Document.Properties.Custom.Add("Producer", "My Producer")

        ' Add a paragraph in which all standard information about the document will be stored.
        Dim builtInPara As New Paragraph(dc, New Run(dc, "Built-in document properties:"), New SpecialCharacter(dc, SpecialCharacterType.LineBreak))
        builtInPara.ParagraphFormat.Alignment = HorizontalAlignment.Left

        For Each docProp In dc.Document.Properties.BuiltIn
            builtInPara.Inlines.Add(New Run(dc, String.Format("{0}: {1}", docProp.Key, docProp.Value)))

            builtInPara.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.LineBreak))
        Next docProp

        ' Add a paragraph in which all additional information about the document will be stored.
        Dim customPropPara As New Paragraph(dc, New Run(dc, "Custom document properties:"), New SpecialCharacter(dc, SpecialCharacterType.LineBreak))
        customPropPara.ParagraphFormat.Alignment = HorizontalAlignment.Left

        For Each docProp In dc.Document.Properties.Custom
            customPropPara.Inlines.Add(New Run(dc, String.Format("{0}: {1} (Type: {2})", docProp.Key, docProp.Value, docProp.Value.GetType())))

            customPropPara.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.LineBreak))
        Next docProp

        ' Add all document properties in the document and save it as PDF file.
        dc.Sections.Clear()
        dc.Sections.Add(New Section(dc, builtInPara, customPropPara))

        dc.Save(statFile, New PdfSaveOptions())

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(statFile) With {.UseShellExecute = True})
    End Sub
End Module