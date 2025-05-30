Imports System.IO
Imports SautinSoft.Document

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			ConvertFromFile()
			ConvertFromStream()
		End Sub
                ''' Get your free trial key here:   
                ''' https://sautinsoft.com/start-for-free/
		''' <summary>
		''' Convert RTF to Text (file to file).
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-rtf-to-text-in-csharp-vb.php
		''' </remarks>
		Private Shared Sub ConvertFromFile()
			Dim inpFile As String = "..\..\..\example.rtf"
			Dim outFile As String = "Result.txt"

			Dim dc As DocumentCore = DocumentCore.Load(inpFile)
			dc.Save(outFile)
			
			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub

		''' <summary>
		''' Convert RTF to Text (using Stream).
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-rtf-to-text-in-csharp-vb.php
		''' </remarks>
		Private Shared Sub ConvertFromStream()

			' We need files only for demonstration purposes.
			' The conversion process will be done completely in memory.
			Dim inpFile As String = "..\..\..\example.rtf"
			Dim outFile As String = "ResultStream.txt"
			Dim inpData() As Byte = File.ReadAllBytes(inpFile)
			Dim outData() As Byte = Nothing

			Using msInp As New MemoryStream(inpData)

				' Load a document.
				Dim dc As DocumentCore = DocumentCore.Load(msInp, New RtfLoadOptions())

				' Save the document to Text format.
				Using outMs As New MemoryStream()
					dc.Save(outMs, New TxtSaveOptions())
					outData = outMs.ToArray()
				End Using
				' Show the result for demonstration purposes.
				If outData IsNot Nothing Then
					File.WriteAllBytes(outFile, outData)
					System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
				End If
			End Using
		End Sub
	End Class
End Namespace
