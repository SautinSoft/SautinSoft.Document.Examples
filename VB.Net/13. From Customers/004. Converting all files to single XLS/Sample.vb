Imports System
Imports System.IO
Imports System.Collections.Generic
Imports SautinSoft.Document
Imports SautinSoft

Namespace Sample
	Friend Class Sample

		Shared Sub Main(ByVal args() As String)
			ConvertToSingleXls()
		End Sub

		''' <summary>
		''' How to convert all files to a single XLS file.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-convert-pdf-docx-rtf-to-single-xls-workbook-net-csharp-vb.php
		''' </remarks>
		Public Shared Sub ConvertToSingleXls()
			' In this example we'll use not only Document .Net component, but also
			' another SautinSoft 'component - PDF Focus .Net (to perform conversion from PDF to single xls workbook).
			' First of all, please perform "Rebuild Solution" to restore PDF Focus .Net package from NuGet.

			' Our steps:
			' 1. Convert all RTF, DOCX, PDF files into a single PDF document. (by Document .Net).
			' 2. Convert the single PDF into a single XLS workbook. (by PDF Focus .Net).

			Dim singlePdfBytes() As Byte = Nothing

			' This file we need only to show intermediate result.
			Dim singlePdfFile As String = "Single.pdf"
			Dim workingDir As String = "..\..\..\"
			Dim singleXlsFile As String = "Single.xls"

			Dim supportedFiles As New List(Of String)()

			For Each file As String In Directory.GetFiles(workingDir, "*.*")
				Dim ext As String = Path.GetExtension(file).ToLower()

				If ext = ".pdf" OrElse ext = ".docx" OrElse ext = ".rtf" Then
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

			' Save our single document into PDF format in memory.
			' Let's save our document to a MemoryStream.
			Using Pdf As New MemoryStream()
				singlePDF.Save(Pdf, New PdfSaveOptions() With {.Compliance = PdfCompliance.PDF_A1a})
				singlePdfBytes = Pdf.ToArray()
			End Using

			' Open the result for demonstration purposes.
			File.WriteAllBytes(singlePdfFile, singlePdfBytes)
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(singlePdfFile) With {.UseShellExecute = True})

			Dim f As SautinSoft.PdfFocus = New PdfFocus()

			f.OpenPdf(singlePdfBytes)

			If f.PageCount > 0 Then
				f.ToExcel(singleXlsFile)
			End If

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(singleXlsFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
