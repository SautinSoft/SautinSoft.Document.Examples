Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			If True Then
				FindWordInParagraph()
			End If
		End Sub
		''' <summary>
		''' Find any "word" in a folder with PDF files inside and show a paragraph, where this word will be found.
		''' You may change the extension: pdf, docx, rtf.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-show-paragraph-containing-required-word-in-csharp-vb-net.php
		''' </remarks>
		Private Shared Sub FindWordInParagraph()
			' A regular expression (shortened as regex or regexp; sometimes referred to as rational expression) is a sequence of characters that specifies a search pattern in text.
			Dim regex As New Regex("\bcompany\b", RegexOptions.IgnoreCase)


			Dim filePath As String = "..\..\..\instance.pdf"

				Dim dc As DocumentCore = DocumentCore.Load(filePath)

				' Provides a functionality to paginate the document content.
				Dim dp As DocumentPaginator = dc.GetPaginator()
				For Each content As ContentRange In dc.Content.Find(regex)
					Dim ef As ElementFrame = dp.GetElementFrames().FirstOrDefault(Function(e) content.Start.Equals(e.Content.Start))
					Dim paragraph As Paragraph = TryCast(content.Start.Parent.Parent, Paragraph)

					' We are looking for a sentence in which this word was found.
					Dim sentence As String = paragraph.Content.ToString().Trim()
					Console.WriteLine("Filename: " & filePath & vbCrLf & sentence)

					' The coordinates of the found word.
					Console.WriteLine("Info:" & ef.Bounds.ToString())
					Console.WriteLine("Next paragraph?")
					Console.ReadKey()
				Next content
		End Sub
	End Class
End Namespace
