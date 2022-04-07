Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        LoadDocFromFile()
        'LoadDocFromStream();
    End Sub

	''' <summary>
	''' Loads a DOC (Word 97-2003) document into DocumentCore (dc) from a file.
	''' </summary>
	''' <remarks>
	''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-doc-word-97-2003-document-net-csharp-vb.php
	''' </remarks>
	Sub LoadDocFromFile()
		Dim filePath As String = "..\..\..\example.doc"
		' The file format is detected automatically from the file extension: ".doc".
		' But as shown in the example below, we can specify DocLoadOptions as 2nd parameter
		' to explicitly set that a loadable document has DOC format.
		Dim dc As DocumentCore = DocumentCore.Load(filePath)

		If dc IsNot Nothing Then
			Console.WriteLine("Loaded successfully!")
		End If

		Console.ReadKey()
	End Sub

	''' <summary>
	''' Loads a DOC (Word 97-2003) document into DocumentCore (dc) from a MemoryStream.
	''' </summary>
	''' <remarks>
	''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-doc-word-97-2003-document-net-csharp-vb.php
	''' </remarks>
	Sub LoadDocFromStream()
		' Assume that we already have a DOC (Word 97-2003) document as bytes array.
		Dim fileBytes() As Byte = File.ReadAllBytes("..\..\..\example.doc")

		Dim dc As DocumentCore = Nothing

		' Create a MemoryStream
		Using ms As New MemoryStream(fileBytes)
			' Load a document from the MemoryStream.
			' Specifying DocLoadOptions we explicitly set that a loadable document is DOC.
			dc = DocumentCore.Load(ms, New DocLoadOptions())
		End Using
		If dc IsNot Nothing Then
			Console.WriteLine("Loaded successfully!")
		End If

		Console.ReadKey()
	End Sub
End Module