Imports System
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables

Namespace Sample
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			FindTextFromTable()
		End Sub

		''' <summary>
		''' How to remove the rows with the specified text from a table.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-find-text-from-table-net-csharp-vb.php
		''' </remarks>
		Public Shared Sub FindTextFromTable()
			Dim longLiverMinYears As Integer = 90

			Dim inpFile As String = "..\..\..\example.docx"
			Dim outFile As String = Path.ChangeExtension(inpFile, ".pdf")

			' Load a document with a table containing various persons with different age.
			Dim dc As DocumentCore = DocumentCore.Load(inpFile)

			' Find a first table in the document.
			Dim table As Table = CType(dc.GetChildElements(True, ElementType.Table).First(), Table)

			' Loop by the all rows from the end.
			' Find long-livers.
			Dim isLongLiver As Boolean = False

			For r As Integer = table.Rows.Count - 1 To 1 Step -1
				isLongLiver = False

				' Take the 3rd cell with the birth date.
				Dim tc As TableCell = table.Rows(r).Cells(2)

				' Get the birth date.
				Dim birthDate As Date = Date.Now
				If Date.TryParse(tc.Content.ToString(), CultureInfo.CreateSpecificCulture("en-US"), DateTimeStyles.None, birthDate) Then
					' Get the person age.
					' Remove the row if the person isn't long-liver.
					If CalculateAge(birthDate) >= longLiverMinYears Then
						isLongLiver = True
					End If
				End If
				' Remove the row if it doesn't contain a long-liver.
				If Not isLongLiver Then
					table.Rows.RemoveAt(r)
				End If
			Next r

			' Save the document as PDF.
			dc.Save(outFile, New PdfSaveOptions())

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
		Private Shared Function CalculateAge(ByVal dateOfBirth As Date) As Integer
			Dim age As Integer = 0
			age = Date.Now.Year - dateOfBirth.Year
			If Date.Now.DayOfYear < dateOfBirth.DayOfYear Then
				age = age - 1
			End If
			Return age
		End Function
	End Class
End Namespace
