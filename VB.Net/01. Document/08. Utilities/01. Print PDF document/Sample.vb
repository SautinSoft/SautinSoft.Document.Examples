Imports Microsoft.VisualBasic
Imports SautinSoft.Document
Imports System
Imports System.IO
Imports System.Text
Imports System.Diagnostics

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			CreateAndPrintPdf()
		End Sub

		''' <summary>
		''' Creates a new document and saves it into PDF format. Print PDF using your default printer.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/print-document-net-csharp-vb.php
		''' </remarks>
		Public Shared Sub CreateAndPrintPdf()
			' Set the path to create pdf document.
			Dim pdfPath As String = "Result.pdf"

			' Let's create a simple PDF document.
			Dim dc As New DocumentCore()

			' Add new section.
			Dim section As New Section(dc)
			dc.Sections.Add(section)

			' Let's set page size A4.
			section.PageSetup.PaperType = PaperType.A4

			' Add a paragraph using ContentRange:
			dc.Content.End.Insert(vbLf & "Hi Dear Friends.", New CharacterFormat() With {
				.Size = 25,
				.FontColor = Color.Blue,
				.Bold = True
			})
			Dim lBr As New SpecialCharacter(dc, SpecialCharacterType.LineBreak)
			dc.Content.End.Insert(lBr.Content)
			dc.Content.End.Insert("I'm happy to see you!", New CharacterFormat() With {
				.Size = 20,
				.FontColor = Color.DarkGreen,
				.UnderlineStyle = UnderlineType.Single
			})

			' Save PDF to a file
			dc.Save(pdfPath, New PdfSaveOptions())

			' Create a new process: Acrobat Reader. You may change in on Foxit Reader.
			Dim processFilename As String = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Microsoft").OpenSubKey("Windows").OpenSubKey("CurrentVersion").OpenSubKey("App Paths").OpenSubKey("Acrobat.exe").GetValue(String.Empty).ToString()

			' Let's transfer our PDF file to the process Adobe Reader
			Dim info As New ProcessStartInfo()
			info.Verb = "print"
			info.FileName = processFilename
			info.Arguments = String.Format("/p /h {0}", pdfPath)
			info.CreateNoWindow = True

			'(It won't be hidden anyway... thanks Adobe!)
			info.WindowStyle = ProcessWindowStyle.Hidden
			info.UseShellExecute = False

			Dim p As Process = Process.Start(info)
			p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden

			' Print our PDF
			Dim counter As Integer = 0
			Do While Not p.HasExited
				System.Threading.Thread.Sleep(1000)
				counter += 1
				If counter = 5 Then
					Exit Do
				End If
			Loop
			If Not p.HasExited Then
				p.CloseMainWindow()
				p.Kill()
			End If

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(pdfPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
