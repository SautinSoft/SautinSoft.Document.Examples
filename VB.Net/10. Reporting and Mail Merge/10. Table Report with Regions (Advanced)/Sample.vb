Imports System
Imports System.Data
Imports System.IO
Imports System.Collections.Generic

Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports System.Globalization

Namespace Sample
	Friend Class Sample
		Shared Sub Main(ByVal args() As String)
			TableReportWithRegionsAdvanced()
		End Sub

		''' <summary>
		''' Generates a table report with regions using XML document as a data source and FieldMerging event.
		''' </summary>
		''' <remarks>
		''' See details at: https://www.sautinsoft.com/products/document/help/net/developer-guide/mail-merge-table-report-with-regions-advanced-net-csharp-vb.php
		''' </remarks>
		Public Shared Sub TableReportWithRegionsAdvanced()
			' Scan directory for image files.
			Dim icons As New Dictionary(Of String, String)()
			For Each iconPath As String In Directory.EnumerateFiles("..\..\..\icons\", "*.jpg")
				icons(Path.GetFileNameWithoutExtension(iconPath.ToLower())) = iconPath

				Select Case Path.GetFileName(iconPath)
					Case "Cherry-apple pie.jpg"
						icons("Cherry/apple/pie".ToLower()) = iconPath
					Case "Dark-milk-white chocolate.jpg"
						icons("Dark/milk/white chocolate".ToLower()) = iconPath
					Case "Rice-lemon-vanilla pudding.jpg"
						icons("Rice/lemon/vanilla pudding".ToLower()) = iconPath
					Case "Spice-cake honey-cake.jpg"
						icons("Spice-cake, honey-cake".ToLower()) = iconPath
					Case "Chocolate-strawberry-vanilla ice cream.jpg"
						icons("Chocolate/strawberry/vanilla ice cream".ToLower()) = iconPath
					Case Else
				End Select
			Next iconPath

			' Create the Dataset and read the XML.
			Dim ds As New DataSet()

			ds.ReadXml("..\..\..\Orders.xml")

			' Load the template document.
			Dim templatePath As String = "..\..\..\InvoiceTemplate.docx"

			Dim dc As DocumentCore = DocumentCore.Load(templatePath)

			' Each product will be decorated by appropriate icon.
			AddHandler dc.MailMerge.FieldMerging, Sub(sender, e)
				' Insert an icon before the product name
				If e.RangeName = "Product" AndAlso e.FieldName = "Name" Then
					e.Inlines.Clear()

					Dim iconPath As String = Nothing
					If icons.TryGetValue(CStr(e.Value).ToLower(), iconPath) Then
						e.Inlines.Add(New Picture(dc, iconPath) With {.Layout = New InlineLayout(New Size(30, 30))})
						e.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.Tab))
					End If

					e.Inlines.Add(New Run(dc, CStr(e.Value)))
					e.Cancel = False
				End If
				' Add the currency sign into "Total" field.
				' You may change the culture "en-GB" to any desired.
				If e.RangeName = "Order" AndAlso e.FieldName = "OrderTotal" Then
					Dim total As Decimal = 0
					If Decimal.TryParse(CStr(e.Value), total) Then
						e.Inlines.Clear()
						e.Inlines.Add(New Run(dc, String.Format(New CultureInfo("en-GB"), "{0:C}", total)))
					End If
				End If
			End Sub

			' Execute the mail merge.
			dc.MailMerge.Execute(ds.Tables("Order"))

			Dim resultPath As String = "Invoices.pdf"

			' Save the output to file
			dc.Save(resultPath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
