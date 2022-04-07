Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        LoadHtmlFromFile()
        'LoadHtmlFromStream();
    End Sub

    ''' <summary>
    ''' Loads an HTML document into DocumentCore (dc) from a file.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-html-document-net-csharp-vb.php
    ''' </remarks>
    Sub LoadHtmlFromFile()
        Dim filePath As String = "..\..\..\example.html"
        ' The file format is detected automatically from the file extension: ".html".
        ' But as shown in the example below, we can specify HtmlLoadOptions as 2nd parameter
        ' to explicitly set that a loadable document has HTML format.
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        If dc IsNot Nothing Then
            Console.WriteLine("Loaded successfully!")
        End If
		
		Console.ReadKey()
    End Sub

    ''' <summary>
    ''' Loads an HTML document into DocumentCore (dc) from a MemoryStream.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-html-document-net-csharp-vb.php
    ''' </remarks>
    Sub LoadHtmlFromStream()
        ' Get document bytes.
        Dim fileBytes() As Byte = File.ReadAllBytes("..\..\..\example.html")

        Dim dc As DocumentCore = Nothing

        ' Create a MemoryStream
        Using ms As New MemoryStream(fileBytes)
            ' Load a document from the MemoryStream.
            ' Specifying HtmlLoadOptions we explicitly set that a loadable document is HTML.
            dc = DocumentCore.Load(ms, New HtmlLoadOptions())
        End Using
        If dc IsNot Nothing Then
            Console.WriteLine("Loaded successfully!")
        End If
		
		Console.ReadKey()		
    End Sub
End Module