Imports System
Imports System.IO
Imports SautinSoft.Document

Namespace Example
	Friend Class Program

		Shared Sub Main(ByVal args() As String)
			LoadPDF()
		End Sub

		''' <summary>
		''' Load a PDF document and get missing fonts.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-pdf-document-get-missing-fonts-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub LoadPDF()
			Dim inpFile As String = "..\..\..\fonts.pdf"
			Dim outFile As String = "Result.docx"

			' Some PDF documents can use rare fonts which isn't installed in your system.
			' During the loading of a PDF document the component tries to find all necessary 
			' fonts installed in your system.
			' In case of the font is missing, you can find its name in the property 'MissingFonts' (after the loading process).
			Dim dc As DocumentCore = DocumentCore.Load(inpFile)

			If SautinSoft.Document.FontSettings.MissingFonts.Count > 0 Then
				Console.WriteLine("Missing Fonts:")
				For Each fontFamily As String In SautinSoft.Document.FontSettings.MissingFonts
					Console.WriteLine(fontFamily)
				Next fontFamily
			End If

			' Next, knowing missing fonts, you can install these fonts in your system.

			' Also, you can specify an extra folder where component should find fonts.
			SautinSoft.Document.FontSettings.UserFontsDirectory = "d:\My Fonts"

			' Furthermore, you can add font substitutes, to use alternative fonts.            
			SautinSoft.Document.FontSettings.AddFontSubstitutes("Melvetika", "Segoe UI")
			Console.WriteLine("We've changed Melvetica font to Segoe UI")

			' Load the document again.
			dc = DocumentCore.Load(inpFile)
			dc.Save(outFile)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
			Console.ReadKey()
		End Sub
	End Class
End Namespace
