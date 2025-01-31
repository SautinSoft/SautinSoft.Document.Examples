Imports SautinSoft.Document

Namespace Example
    Class Program
        Shared Sub Main(ByVal args As String())
            LoadPDF()
        End Sub

        Shared Sub LoadPDF()
            Dim inpFile As String = "..\..\..\fonts.docx"
            Dim outFile As String = "Result.pdf"
            Dim dc As DocumentCore = DocumentCore.Load(inpFile)
            Dim MissingFonts As List(Of String) = New List(Of String)()
            AddHandler FontSettings.FontSelection, Sub(s, e)
                                                       ' Search for embedded and missing fonts.
                                                       If Not MissingFonts.Contains(e.FontName) Then
                                                           MissingFonts.Add(e.FontName)
                                                       End If

                                                       ' Replacing the Dubai font with Segoe UI.
                                                       Dim segoe = FontSettings.Fonts.FirstOrDefault(Function(f) f.FamilyName = "Segoe UI")
                                                       If e.FontName = "Dubai" Then
                                                           e.SelectedFont = segoe
                                                           Console.WriteLine("We've changed Dubai font to Segoe UI")
                                                       End If
                                                   End Sub

            dc.Save(outFile)

            If MissingFonts.Count > 0 Then
                Console.WriteLine("Missing Fonts:")

                For Each fontFamily As String In MissingFonts
                    Console.WriteLine(fontFamily)
                Next
            End If

            FontSettings.FontsBaseDirectory = "d:\My Fonts"
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {
                .UseShellExecute = True
            })
        End Sub
    End Class
End Namespace
