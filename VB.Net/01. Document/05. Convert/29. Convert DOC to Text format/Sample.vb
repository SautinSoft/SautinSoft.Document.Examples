Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        ConvertFromFile()
        ConvertFromStream()
    End Sub

	''' <summary>
	''' Convert DOC (Word 97-2003) to Text (file to file).
	''' </summary>
	''' <remarks>
	''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-doc-word-97-2003-to-text-in-csharp-vb.php
	''' </remarks>
	Sub ConvertFromFile()
		Dim inpFile As String = "..\..\..\example.doc"
		Dim outFile As String = "Result.txt"

		Dim dc As DocumentCore = DocumentCore.Load(inpFile)
		dc.Save(outFile)

		' Open the result for demonstration purposes.
		System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
	End Sub

	''' <summary>
	''' Convert DOC (Word 97-2003) to Text (using Stream).
	''' </summary>
	''' <remarks>
	''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-doc-word-97-2003-to-text-in-csharp-vb.php
	''' </remarks>
	Sub ConvertFromStream()

		' We need files only for demonstration purposes.
		' The conversion process will be done completely in memory.
		Dim inpFile As String = "..\..\..\example.doc"
		Dim outFile As String = "ResultStream.txt"
		Dim inpData() As Byte = File.ReadAllBytes(inpFile)
		Dim outData() As Byte = Nothing

		Using msInp As New MemoryStream(inpData)

			' Load a document.
			Dim dc As DocumentCore = DocumentCore.Load(msInp, New DocLoadOptions())

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
End Module