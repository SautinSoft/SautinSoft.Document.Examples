Imports SautinSoft.Document
Imports System
Imports System.IO
Imports System.Linq
Imports System.Reflection.Metadata

Namespace Sample
	Friend Class Sample
		Shared Sub Main(ByVal args() As String)
			' Get your free trial key here:   
			' https://sautinsoft.com/start-for-free/

			ConvertRTLcontent()
		End Sub

		''' <summary>
		''' How to convert documents with Right-To-Left content to HTML.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-documents-with-right-to-left-content-to-html.php
		''' </remarks>
		Public Shared Sub ConvertRTLcontent()
			Dim sourcePath As String = "..\..\..\RTL.docx"
			Dim destPath As String = "RTL.html"

			' Load document with arabic, hindi, hebrew content.
			Dim dc As DocumentCore = DocumentCore.Load(sourcePath)

			' Save the document as HTML.
			dc.Save(destPath, New HtmlFixedSaveOptions())

			' Show the source and the dest documents.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(sourcePath) With {.UseShellExecute = True})
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(destPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
