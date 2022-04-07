Imports System.Text
Imports System.Linq
Imports SautinSoft.Document

Namespace Sample
	Friend Class Sample
		Shared Sub Main(ByVal args() As String)
			DeleteHyperlinksObjects()
			DeleteHyperlinksURL()
		End Sub

		''' <summary>
		''' How to delete all hyperlink objects.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/hyperlinks-delete-url-csharp-vb-net.php
		''' </remarks>		
		Public Shared Sub DeleteHyperlinksObjects()
			' Let us say, we've a DOCX document.
			' And we've to remove the hyperlink objects.

			Dim inpFile As String = "..\..\..\Hyperlinks example.docx"
			Dim outFile As String = "Result - Delete Hyperlinks completely.pdf"

			' Let's open our document.
			Dim dc As DocumentCore = DocumentCore.Load(inpFile)

			' Specify the custom URL.
			Dim customURL As String = "https://www.sautinsoft.com"

			' Loop by all hyperlinks and replace the URL (address).
			For Each hpl As Hyperlink In dc.GetChildElements(True, ElementType.Hyperlink).Reverse()
				hpl.ParentCollection.Remove(hpl)
			Next hpl

			' Save our document back, but in PDF format.
			dc.Save(outFile, New PdfSaveOptions() With {.Compliance = PdfCompliance.PDF_14})

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
		''' <summary>
		''' How to delete all hyperlinks and leave only their text.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/hyperlinks-delete-url-csharp-vb-net.php
		''' </remarks>		
		Public Shared Sub DeleteHyperlinksURL()
			' Let us say, we've a DOCX document.
			' And we've to remove all hyperlinks, leave only their text.

			' Note, we can't make the property 'Hyperlink.Address' empty, this is not allowed.
			' Therefore we have to remove the all 'Hyperlinks' object and 
			' insert the text objects 'Inline' instead of them.

			Dim inpFile As String = "..\..\..\Hyperlinks example.docx"
			Dim outFile As String = "Result - delete links and leave text.docx"

			' Let's open our document.
			Dim dc As DocumentCore = DocumentCore.Load(inpFile)

			' Loop by all hyperlinks in a reverse, to remove the "Hyperlink" objects 
			' and replace them by their text ("Inline" objects).
			For Each hpl As Hyperlink In dc.GetChildElements(True, ElementType.Hyperlink).Reverse()
				' Get the "Hyperlink" index in the parent collection.
				Dim parentCollection As InlineCollection = hpl.ParentCollection
				Dim index As Integer = parentCollection.IndexOf(hpl)

				' Get the "Hyperlink" text as the Inline collection.
				Dim textInlines As InlineCollection = hpl.DisplayInlines

				' Remove the "Hyperlink" object from the parent collection by index.
				parentCollection.RemoveAt(index)

				' Insert the text (collection of Inlines) instead of the removed "Hyperlink" object 
				' into the parent collection.
				For i As Integer = 0 To textInlines.Count - 1
					' Set the Auto font color (Black for the most cases) and remove the underline.
					' Hide these lines if you want to leave the formatting the same as the hyperlink had.
					If TypeOf textInlines(i) Is Run Then
						TryCast(textInlines(i), Run).CharacterFormat.FontColor = Color.Red
						TryCast(textInlines(i), Run).CharacterFormat.UnderlineStyle = UnderlineType.None
					End If
					parentCollection.Insert(index + i, textInlines(i).Clone(True))
				Next i
			Next hpl
			' Save the document back.
			dc.Save(outFile)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
