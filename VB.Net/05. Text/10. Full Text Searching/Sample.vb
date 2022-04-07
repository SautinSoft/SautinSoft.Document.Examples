Imports System
Imports System.IO
Imports System.Collections.Generic
Imports SautinSoft.Document
Imports System.Drawing
Imports System.Linq
Imports System.Text.RegularExpressions



Namespace Sample
    Friend Class Sample

        Shared Sub Main(ByVal args() As String)
            Dim searchDir As String = Path.GetFullPath("..\..\..\searching\")
            Dim searchText As String = "with"
            FullTextSearching(searchDir, searchText)
        End Sub

        ''' <summary>
        ''' This sample shows how to launch full text search in the specific directory.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/full-text-searching-in-documents-net-csharp-vb.php
        ''' </remarks>
        Public Shared Sub FullTextSearching(ByVal searchPath As String, ByVal searchText As String)
            Dim searchDir As New DirectoryInfo(searchPath)
            Dim supportedFiles As New List(Of String)()

            ' 1. Find theS files to make search.
            ' Specify to make the search only in *.docx, *.rtf, *.pdf and *.html files,
            ' including subdirectories.
            For Each file As String In Directory.GetFiles(searchDir.FullName, "*.*", SearchOption.AllDirectories)
                Dim ext As String = Path.GetExtension(file).ToLower()

                If ext = ".docx" OrElse ext = ".pdf" OrElse ext = ".html" OrElse ext = ".rtf" Then
                    supportedFiles.Add(file)
                End If
            Next file

            ' 2. Perform the text search in the each file using a loop.
            ' We'll search the word "video" in the each and count how many times the file contains it.
            Console.WriteLine($"The results for ""{searchText}"":")

            Dim totalFiles As Integer = 0, totalMatches As Integer = 0
            For Each file As String In supportedFiles
                Dim dc As DocumentCore = DocumentCore.Load(file)
                totalFiles += 1
                Dim regex As New Regex($"\b({searchText})\b", RegexOptions.IgnoreCase)

                ' Show also subfolder if we aren't in the root folder.
                Dim dirInfo As New DirectoryInfo(Path.GetDirectoryName(file))
                Dim fileName As String = String.Empty

                If dirInfo.FullName.TrimEnd(New Char() {"\"c}) <> searchDir.FullName.TrimEnd(New Char() {"\"c}) Then
                    fileName = file.Substring(searchPath.Length, file.Length - searchPath.Length)
                Else
                    ' We are in the root folder.
                    fileName = Path.GetFileName(file)
                End If

                Dim matches As Integer = dc.Content.Find(regex).Count()
                totalMatches += matches

                Console.WriteLine($"{totalFiles:D3} from {supportedFiles.Count} {fileName} - {matches} matches.")
            Next file
            Console.WriteLine($"Searching finished. {supportedFiles.Count} file(s) has been processed. Total matches: {totalMatches}.")
            Console.WriteLine("Press any key ...")
            Console.ReadKey()
        End Sub
    End Class
End Namespace