Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        LoadDocxFromFile()
        'LoadDocxFromStream();
    End Sub

    ''' <summary>
    ''' Loads a DOCX document into DocumentCore (dc) from a file.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-docx-document-net-csharp-vb.php
    ''' </remarks>
    Sub LoadDocxFromFile()
        Dim filePath As String = "..\..\..\example.docx"
        ' The file format is detected automatically from the file extension: ".docx".
        ' But as shown in the example below, we can specify DocxLoadOptions as 2nd parameter
        ' to explicitly set that a loadable document has Docx format.
        Dim dc As DocumentCore = DocumentCore.Load(filePath)

        If dc IsNot Nothing Then
            Console.WriteLine("Loaded successfully!")
        End If
		
		Console.ReadKey()
    End Sub

    ''' <summary>
    ''' Loads a DOCX document into DocumentCore (dc) from a MemoryStream.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-docx-document-net-csharp-vb.php
    ''' </remarks>
    Sub LoadDocxFromStream()
        ' Assume that we already have a DOCX document as bytes array.
        Dim fileBytes() As Byte = File.ReadAllBytes("..\..\..\example.docx")

        Dim dc As DocumentCore = Nothing

        ' Create a MemoryStream
        Using ms As New MemoryStream(fileBytes)
            ' Load a document from the MemoryStream.
            ' Specifying DocxLoadOptions we explicitly set that a loadable document is Docx.
            dc = DocumentCore.Load(ms, New DocxLoadOptions())
        End Using
        If dc IsNot Nothing Then
            Console.WriteLine("Loaded successfully!")
        End If
		
		Console.ReadKey()		
    End Sub
End Module