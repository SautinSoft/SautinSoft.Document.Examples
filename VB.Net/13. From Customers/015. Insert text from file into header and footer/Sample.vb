Imports System
Imports System.IO
Imports System.Collections.Generic
Imports SautinSoft.Document
Imports System.Linq
Imports SautinSoft.Document.Tables
Imports System.Text

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertTextHeaderFooter()
		End Sub

		''' <summary>
		''' How to insert the contents of a file into the header and footer of an existing HTML.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-insert-text-from-file-into-header-footer-in-csharp-vb-net.php
		''' </remarks>

		Private Shared Sub InsertTextHeaderFooter()
			Dim inpFile As String = "..\..\..\example.html"
			Dim outFile As String = "result.docx"

			' Reads all text from HTML file.
			Dim htmlHeaderBytes() As Byte = Encoding.UTF8.GetBytes(File.ReadAllText("..\..\..\header.html"))
			Dim htmlHeader As String = System.Text.Encoding.UTF8.GetString(htmlHeaderBytes)

			' Reads all text from RTF file.
			Dim rtfFooterBytes() As Byte = Encoding.UTF8.GetBytes(File.ReadAllText("..\..\..\footer.rtf"))
			Dim rtfFooter As String = System.Text.Encoding.UTF8.GetString(rtfFooterBytes)

			' Load a document.
			Dim dc As DocumentCore = DocumentCore.Load(inpFile)

			' Create a new header with formatted HTML text.
			Dim header As New HeaderFooter(dc, HeaderFooterType.HeaderDefault)

			' Add the header into HeadersFooters collection and Clone to all sections.
			header.Content.Start.Insert(htmlHeader, LoadOptions.HtmlDefault)
			For Each s As Section In dc.Sections
					s.HeadersFooters.Add(header.Clone(True))
			Next s

			' Create a new footer with formatted RTF text.
			Dim footer As New HeaderFooter(dc, HeaderFooterType.FooterDefault)

			' Add the footer into HeadersFooters collection and Clone to all sections.
			footer.Content.Start.Insert(rtfFooter, LoadOptions.RtfDefault)
			For Each s As Section In dc.Sections
				s.HeadersFooters.Add(footer.Clone(True))
			Next s

			' Save the result as DOCX file.
			dc.Save(outFile)

				' Open the result for demonstration purposes.
				System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace