Imports System
Imports System.IO
Imports System.Collections.Generic
Imports SautinSoft.Document

Namespace Sample
	Friend Class Sample

		Shared Sub Main(ByVal args() As String)
			MergeDocumentsInMem()
		End Sub

		''' <summary>
		''' This sample shows how to merge multiple files DOCX, PDF into single document in memory.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/merge-multiple-documents-docx-pdf-in-memory-net-csharp-vb.php
		''' </remarks>
		Public Shared Sub MergeDocumentsInMem()
			' We'll use these files only to get data and show the result.
			' The whole merging process will be done in memory.
			Dim documents() As String = {"..\..\..\one.docx", "..\..\..\two.pdf"}
			Dim singleDocumentPath As String = "Result.pdf"

			' Read the files and retrieve the file data into this Dictionary.
			' Thus we'll have input data completely in memory.
			Dim documentsData As New Dictionary(Of String, Byte())()
			For Each file As String In documents
				documentsData.Add(file, System.IO.File.ReadAllBytes(file))
			Next file

			' Merge documents in memory (using MemoryStream)
			' 1. Create a single document.
			Dim dcSingle As New DocumentCore()

			For Each document As KeyValuePair(Of String, Byte()) In documentsData
				' Create new MemoryStream based on document byte array.
				Using msDoc As New MemoryStream(document.Value)
					Dim lo As LoadOptions = Nothing
					If Path.GetExtension(document.Key).ToLower() = ".docx" Then
						lo = New DocxLoadOptions()
					ElseIf Path.GetExtension(document.Key).ToLower() = ".pdf" Then
						lo = New PdfLoadOptions()
					End If

					Dim dc As DocumentCore = DocumentCore.Load(msDoc, lo)

					Console.WriteLine("Adding: {0}...", Path.GetFileName(document.Key))

					' Create import session.
					Dim session As New ImportSession(dc, dcSingle, StyleImportingMode.KeepSourceFormatting)

					' Loop through all sections in the source document.
					For Each sourceSection As Section In dc.Sections
						' Because we are copying a section from one document to another,
						' it is required to import the Section into the destination document.
						' This adjusts any document-specific references to styles, bookmarks, etc.
						'
						' Importing a element creates a copy of the original element, but the copy
						' is ready to be inserted into the destination document.
						Dim importedSection As Section = dcSingle.Import(Of Section)(sourceSection, True, session)

						' First section start from new page.
						If dc.Sections.IndexOf(sourceSection) = 0 Then
							importedSection.PageSetup.SectionStart = SectionStart.NewPage
						End If

						' Now the new section can be appended to the destination document.
						dcSingle.Sections.Add(importedSection)
					Next sourceSection
				End Using
			Next document

			' Save the resulting document as PDF into MemoryStream.
			Using msPdf As New MemoryStream()
				dcSingle.Save(msPdf, New PdfSaveOptions())

				' Let's also save our document to PDF file for showing the result.
				File.WriteAllBytes(singleDocumentPath, msPdf.ToArray())

				' Open the result for demonstration purposes.
				System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(singleDocumentPath) With {.UseShellExecute = True})
			End Using
		End Sub
	End Class
End Namespace
