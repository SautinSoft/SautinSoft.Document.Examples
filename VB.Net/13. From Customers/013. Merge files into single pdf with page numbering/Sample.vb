Option Infer On
Imports System
Imports System.IO
Imports System.Collections.Generic
Imports SautinSoft.Document
Imports System.Linq

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			MergeMultipleDocuments()
		End Sub

		''' <summary>
		''' This sample shows how to merge DOCX, RTF, PDF, PNG or text files into a single PDF document and add page numbers inside.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-merge-multiple-files-into-single-pdf-add-page-numbering-in-csharp-vb-net.php
		''' </remarks>
		Public Shared Sub MergeMultipleDocuments()
			' Path to our combined document.
			Dim singlePDFPath As String = "single.pdf"

			Dim supportedFiles As New List(Of String)()

			' Sort files by name. This way the files will be combined in alphabetical order a-, b-, c- or 1-, 2-, 3-...
			Dim files = Directory.GetFiles("..\..\..\DirToMerge\", "*").OrderByDescending(Function(d) (New FileInfo(d)).Name)

			' Fill the collection 'supportedFiles' by *.docx, *.pdf, *.png and *.txt.
			For Each file As String In files
				Dim ext As String = Path.GetExtension(file)

				If ext = ".docx" OrElse ext = ".pdf" OrElse ext = ".txt" OrElse ext = ".png" Then
					supportedFiles.Add(file)
				End If
			Next file

			' Create single pdf.
			Dim singlePDF As New DocumentCore()

			For Each file As String In supportedFiles
				Dim dc As DocumentCore = DocumentCore.Load(file)

				Console.WriteLine("Adding: {0}...", Path.GetFileName(file))

				' Create import session.
				Dim session As New ImportSession(dc, singlePDF, StyleImportingMode.KeepSourceFormatting)

				' Loop through all sections in the source document.
				For Each sourceSection As Section In dc.Sections
					' Because we are copying a section from one document to another,
					' it is required to import the Section into the destination document.
					' This adjusts any document-specific references to styles, bookmarks, etc.
					'
					' Importing a element creates a copy of the original element, but the copy
					' is ready to be inserted into the destination document.
					Dim importedSection As Section = singlePDF.Import(Of Section)(sourceSection, True, session)

					' First section start from new page.
					If dc.Sections.IndexOf(sourceSection) = 0 Then
						importedSection.PageSetup.SectionStart = SectionStart.NewPage
					End If

					' Now the new section can be appended to the destination document.
					singlePDF.Sections.Add(importedSection)
				Next sourceSection
			Next file
			' We place our page numbers into the footer.
			' Therefore we've to create a footer.
			Dim footer As New HeaderFooter(singlePDF, HeaderFooterType.FooterDefault)

			' Create a new paragraph to insert a page numbering.
			' So that, our page numbering looks as: Page N of M.
			Dim par As New Paragraph(singlePDF)
			par.ParagraphFormat.Alignment = HorizontalAlignment.Right
			Dim cf As New CharacterFormat() With {
				.FontName = "Consolas",
				.Size = 18.0,
				.FontColor = Color.Red
			}
			par.Content.Start.Insert("Page ", cf.Clone())

			' Page numbering is a Field.
			Dim fPage As New Field(singlePDF, FieldType.Page)
			fPage.CharacterFormat = cf.Clone()
			par.Content.End.Insert(fPage.Content)

			footer.Blocks.Add(par)

			For Each s As Section In singlePDF.Sections
				s.HeadersFooters.Add(footer.Clone(True))
			Next s

			' Save single PDF to a file.
			singlePDF.Save(singlePDFPath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(singlePDFPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace