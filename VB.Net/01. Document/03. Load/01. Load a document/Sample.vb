Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        LoadFromFile()
        'LoadFromStream()
        'LoadFromBytes()
    End Sub

    ''' <summary>
    ''' Loads a document into DocumentCore (dc) from a file.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-document.php
    ''' </remarks>
    Sub LoadFromFile()
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
    ''' Loads a document into DocumentCore (dc) from a Stream.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-document.php
    ''' </remarks>
    Sub LoadFromStream()
        ' We've knowingly created an empty DocumentCore instance before "Using {}"
        ' to continue work with it after stream will be closed.
        Dim dc As DocumentCore = Nothing
        Using fs As New FileStream("..\..\..\example.docx", FileMode.Open)

            ' Here we explicitly set that a loadable document is Docx.
            dc = DocumentCore.Load(fs, New DocxLoadOptions())
        End Using
        If dc IsNot Nothing Then
            Console.WriteLine("Loaded successfully!")
        End If

        Console.ReadKey()
    End Sub

    ''' <summary>
    ''' Loads a document into DocumentCore (dc) from an array of bytes.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-document.php
    ''' </remarks>
    Sub LoadFromBytes()
        ' Get document bytes from a file.
        Dim fileBytes() As Byte = File.ReadAllBytes("..\..\..\example.pdf")

        Dim dc As DocumentCore = Nothing
        Using ms As New MemoryStream(fileBytes)

            ' With PdfLoadOptions we explicitly set that a loadable document is PDF.
            Dim pdfLO As New PdfLoadOptions()
            With pdfLO
                .RasterizeVectorGraphics = False
                .DetectTables = False
                ' 'Disabled' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
                ' 'Enabled' - Always load embedded fonts in PDF.
                ' 'Auto' - Load only embedded fonts missing in the system. In other case, use the system fonts.			
                .PreserveEmbeddedFonts = PropertyState.Auto
                .PageIndex = 0
                .PageCount = 2
            End With

            ' RasterizeVectorGraphics = False
            ' This means to load vector graphics as is. Don't transform it to raster images.

            ' DetectTables = False
            ' This means don't detect tables.
            ' The PDF format doesn't have real tables, in fact it's a set of orthogonal graphic lines.
            ' Set it to 'True' and the component will detect and recreate tables from graphic lines.

            dc = DocumentCore.Load(ms, pdfLO)
        End Using
        If dc IsNot Nothing Then
            Console.WriteLine("Loaded successfully!")
        End If

        Console.ReadKey()
    End Sub
End Module