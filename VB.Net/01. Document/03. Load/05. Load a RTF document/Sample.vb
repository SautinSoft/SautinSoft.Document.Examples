Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        'LoadRtfFromStream();
        LoadRtfFromFile()
    End Sub

    ''' <summary>
    ''' Loads an RTF document into DocumentCore (dc) from a file.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-rtf-document-net-csharp-vb.php
    ''' </remarks>
    Sub LoadRtfFromFile()
        Dim filePath As String = "..\..\..\example.rtf"

        ' The file format is detected automatically from the file extension: ".rtf".
        ' But as shown in the example below, we can specify RtfLoadOptions as 2nd parameter
        ' to explicitly set that a loadable document has RTF format.
        Dim dc As DocumentCore = DocumentCore.Load(filePath)

        If dc IsNot Nothing Then
            Console.WriteLine("Loaded successfully!")
        End If
		
		Console.ReadKey()		
    End Sub

    ''' <summary>
    ''' Loads an RTF document into DocumentCore (dc) from a file.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-rtf-document-net-csharp-vb.php
    ''' </remarks>
    Sub LoadRtfFromStream()
        ' Get document bytes.
        Dim fileBytes() As Byte = File.ReadAllBytes("..\..\..\example.rtf")

        Dim dc As DocumentCore = Nothing

        ' Create a MemoryStream
        Using ms As New MemoryStream(fileBytes)
            ' Load a document from the MemoryStream.
            ' Specifying RtfLoadOptions we explicitly set that a loadable document is RTF.
            dc = DocumentCore.Load(ms, New RtfLoadOptions())
        End Using
        If dc IsNot Nothing Then
            Console.WriteLine("Loaded successfully!")
        End If
		
		Console.ReadKey()		
    End Sub

End Module