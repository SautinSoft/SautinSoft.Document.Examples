Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        LoadPDFFromFile()
        'LoadPDFFromStream()
    End Sub

    ''' <summary>
    ''' Loads a PDF document into DocumentCore (dc) from a file.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-pdf-document-net-csharp-vb.php
    ''' </remarks>
    Sub LoadPDFFromFile()
        Dim filePath As String = "..\..\..\example.pdf"

        ' The file format is detected automatically from the file extension: ".pdf".
        ' But as shown in the example below, we can specify PdfLoadOptions as 2nd parameter
        ' to explicitly set that a loadable document has PDF format.
        Dim dc As DocumentCore = DocumentCore.Load(filePath)

        If dc IsNot Nothing Then
            Console.WriteLine("Loaded successfully!")
        End If

        Console.ReadKey()
    End Sub

    ''' <summary>
    ''' Loads a PDF document into DocumentCore (dc) from a MemoryStream.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-pdf-document-net-csharp-vb.php
    ''' </remarks>
    Sub LoadPDFFromStream()
        ' Assume that we already have a PDF document as bytes array.
        Dim fileBytes() As Byte = File.ReadAllBytes("..\..\..\example.pdf")

        Dim dc As DocumentCore = Nothing

        ' Create a MemoryStream
        Using pdfStream As New MemoryStream(fileBytes)

            ' Specifying PdfLoadOptions we explicitly set that a loadable document is PDF.
            Dim pdfLO As New PdfLoadOptions()
            With pdfLO
                .RasterizeVectorGraphics = False
                .DetectTables = False
                ' 'Disabled' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
                ' 'Enabled' - Always load embedded fonts in PDF.
                ' 'Auto' - Load only embedded fonts missing in the system. In other case, use the system fonts.
                .PreserveEmbeddedFonts = PropertyState.Auto
                .PageIndex = 0
                .PageCount = 1
            End With

            ' RasterizeVectorGraphics = False
            ' This means to load vector graphics as is. Don't transform it to raster images.

            ' DetectTables = False
            ' This means don't detect tables.
            ' The PDF format doesn't have real tables, in fact it's a set of orthogonal graphic lines.
            ' Set it to 'True' and the component will detect and recreate tables from graphic lines.

            ' Load a PDF document from the MemoryStream.
            dc = DocumentCore.Load(pdfStream, New PdfLoadOptions())
        End Using
        If dc IsNot Nothing Then
            Console.WriteLine("Loaded successfully!")
        End If

        Console.ReadKey()
    End Sub
End Module