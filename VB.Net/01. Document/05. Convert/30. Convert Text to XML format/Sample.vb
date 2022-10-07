Imports System.IO
Imports SautinSoft.Document

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			ConvertFromFile()
		End Sub

		''' <summary>
		''' Convert Text to XML (file to file).
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-text-to-xml-in-csharp-vb.php
		''' </remarks>
		Private Shared Sub ConvertFromFile()
			Dim inpFile As String = "..\..\..\example.txt"
			Dim outFile As String = "result.xml"

			Dim dc As DocumentCore = DocumentCore.Load(inpFile)

			' Convert non-tabular data and saves it as Xml file.
			dc.Save(outFile, New XmlSaveOptions With {.ConvertNonTabularDataToSpreadsheet = True})

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace