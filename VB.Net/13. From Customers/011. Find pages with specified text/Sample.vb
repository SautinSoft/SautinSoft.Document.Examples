Option Infer On

Imports SautinSoft.Document
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Xml.Linq

Namespace Sample
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			FindPagesSpecifiedText()
		End Sub
                ''' Get your free 30-day key here:   
                ''' https://sautinsoft.com/start-for-free/
		''' <summary>
		''' How to find out on which pages of the document the required word is located.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-find-pages-with-specified-text-net-csharp-vb.php
		''' </remarks>
		Public Shared Sub FindPagesSpecifiedText()
			' The path for input files or directory.
			Dim inpFile As String = "..\..\..\example.docx"

			' What we need to search.
			Dim searchText = "Invoice"
			Dim quantity As Integer = 0

			' Load our documument in Document's engine.
			Dim dc As DocumentCore = DocumentCore.Load(inpFile)

			' Regex https://en.wikipedia.org/wiki/Regular_expression
			Dim regex As New Regex(searchText, RegexOptions.IgnoreCase)

			' Document paginator allows you to calculate of pages.
			Dim dp As DocumentPaginator = dc.GetPaginator()

			' We will search "searchText" on each pages (enumeration).
			For page As Integer = 0 To dp.Pages.Count - 1
				For Each item As ContentRange In dp.Pages(page).Content.Find(regex).Reverse()
					Console.WriteLine($"I see the [{searchText}] on the page # {page + 1}")
					quantity += 1
				Next item
			Next page
			Console.WriteLine()
			Console.WriteLine($"I met [{searchText}] {quantity} times.  Please click on any button")
			Console.ReadKey()
		End Sub
	End Class
End Namespace