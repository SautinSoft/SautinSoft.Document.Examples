Imports System.IO
Imports SautinSoft.Document

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			ConvertFromFile()
			ConvertFromStream()
		End Sub

		''' <summary>
		''' Convert DOCX to Text (file to file).
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-docx-to-text-in-csharp-vb.php
		''' </remarks>
		Private Shared Sub ConvertFromFile()
			Dim inpFile As String = "..\..\..\example.docx"
			Dim outFile As String = "Result.txt"

			Dim dc As DocumentCore = DocumentCore.Load(inpFile)
			dc.Save(outFile)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub

		''' <summary>
		''' Convert DOCX to Text (using Stream).
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-docx-to-text-in-csharp-vb.php
		''' </remarks>
		Private Shared Sub ConvertFromStream()

			' We need files only for demonstration purposes.
			' The conversion process will be done completely in memory.
			Dim inpFile As String = "..\..\..\example.docx"
			Dim outFile As String = "ResultStream.txt"
			Dim inpData() As Byte = File.ReadAllBytes(inpFile)
			Dim outData() As Byte = Nothing

			Using msInp As New MemoryStream(inpData)

				' Load a document.
				Dim dc As DocumentCore = DocumentCore.Load(msInp, New DocxLoadOptions())

				' Save the document to text format.
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
