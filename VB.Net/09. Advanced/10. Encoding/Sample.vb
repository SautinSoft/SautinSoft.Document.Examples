Imports SautinSoft.Document
Imports System.IO
Imports System.Text

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free trial key here:   
			' https://sautinsoft.com/start-for-free/

			Unicode()
		End Sub
		''' <summary>
		''' Convert files using different encodings.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/encodings.php
		''' </remarks>
		Private Shared Sub Unicode()
			Dim inpFile As String = "..\..\..\example.txt"
			Dim outFile As String = "Result.pdf"

			' Provides access to an encoding provider for code pages that otherwise are available only in the desktop .NET Framework and .NET
			' https://learn.microsoft.com/en-us/dotnet/api/system.text.codepagesencodingprovider
			System.Text.Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)

			Dim dc As DocumentCore = DocumentCore.Load(inpFile, New TxtLoadOptions With {.Encoding = System.Text.Encoding.GetEncoding(1251)})
			dc.Save(outFile)

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
