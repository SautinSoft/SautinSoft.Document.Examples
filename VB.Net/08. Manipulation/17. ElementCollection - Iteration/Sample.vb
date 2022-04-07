Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        ShowInlines()
    End Sub

    ''' <summary>
    ''' Iterates through a document and count the amount of Paragraphs and Runs.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/elementcollection-iteration.php
    ''' </remarks>
    Sub ShowInlines()
        Dim filePath As String = "..\..\..\example.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        Console.WriteLine("This document contains from:")
        For sect As Integer = 0 To dc.Sections.Count - 1
            Console.WriteLine("Section {0} contains from:", sect)
            Dim totalParagraphs As Integer = 0
            Dim section As Section = dc.Sections(sect)
            For blocks As Integer = 0 To section.Blocks.Count - 1
                If TypeOf section.Blocks(blocks) Is Paragraph Then
                    totalParagraphs += 1
                    Dim paragraph As Paragraph = TryCast(section.Blocks(blocks), Paragraph)
                    Console.Write(vbTab & vbTab & " Paragraph {0} contains from: ", totalParagraphs)
                    Dim totalRuns As Integer = 0
                    For i As Integer = 0 To paragraph.Inlines.Count - 1
                        If TypeOf paragraph.Inlines(i) Is Run Then
                            totalRuns += 1
                        End If
                    Next i
                    Console.WriteLine("{0} Run(s).", totalRuns)
                    Console.ReadKey()
                End If
            Next blocks
        Next sect
    End Sub
End Module