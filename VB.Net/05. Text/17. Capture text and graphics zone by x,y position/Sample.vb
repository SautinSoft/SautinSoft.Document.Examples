Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports SautinSoft.Document.Tables

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free 30-day key here:   
			' https://sautinsoft.com/start-for-free/

			CaptureTextZoneByXY()
		End Sub
		''' <summary>
		''' How to capture text and images from the existing PDF, DOCX, any document by specific (x,y) coordinates
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/capture-text-images-in-pdf-docx-document-by-specific-zone-x-y-coordinates-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub CaptureTextZoneByXY()
			' Let us say, we want to capture text and graphics into:            
			' Left-Top:(0, 50) mm,
			Dim mmXY = (0F, 50.0F)
			' Width: 250 mm, Height: 150 mm.
			Dim mmWH = (250.0F, 150.0F)

			' Zero-page index, e.g. page 1 has index 0.
			Dim pageCollection() As Integer = {0}

			' Convert mm to points
			Dim leftX As Double = LengthUnitConverter.Convert(mmXY.Item1, LengthUnit.Millimeter, LengthUnit.Point)
			Dim topY As Double = LengthUnitConverter.Convert(mmXY.Item2, LengthUnit.Millimeter, LengthUnit.Point)
			Dim width As Double = LengthUnitConverter.Convert(mmWH.Item1, LengthUnit.Millimeter, LengthUnit.Point)
			Dim height As Double = LengthUnitConverter.Convert(mmWH.Item2, LengthUnit.Millimeter, LengthUnit.Point)

			Dim inpFile As String = Path.GetFullPath("..\..\..\Potato Beetle.pdf")
			Dim outFile As String = "Result.docx"

			' 1. Load an existing document, load only specigic pages.
			Dim opt As New PdfLoadOptions() With {
				.SelectedPages = pageCollection,
				.DetectTables = True,
				.ConversionMode = PdfConversionMode.Flowing
			}
			Dim dc As DocumentCore = DocumentCore.Load(inpFile, opt)

			' 2. Create new document to store captured data.
			Dim dcCaptured As New DocumentCore()
			' Create import session.
			Dim session As New ImportSession(dc, dcCaptured, StyleImportingMode.KeepSourceFormatting)

			' 3. Iterate through document pages
			' and capture elements by (X,Y) and (width, height).
			Dim paginator = dc.GetPaginator(New PaginatorOptions() With {.UpdateFields = True})
			Dim pageIndex As Integer = 0
			For Each page In paginator.Pages
				Dim importedSection As Section = Nothing
				For Each elementFrame In page.GetElementFrames(ElementType.Paragraph, ElementType.Table)
					' Is element inside capturing zone?
					If elementFrame.Bounds.Left >= leftX AndAlso elementFrame.Bounds.Left <= leftX + width AndAlso elementFrame.Bounds.Top >= topY AndAlso elementFrame.Bounds.Top <= topY + height Then

						If importedSection Is Nothing Then
							importedSection = dcCaptured.Import(Of Section)(dc.Sections(pageIndex), False, session)
							dcCaptured.Sections.Add(importedSection)
						End If

						Dim tempVar As Boolean = TypeOf elementFrame.Element Is Paragraph
						Dim par As Paragraph = If(tempVar, CType(elementFrame.Element, Paragraph), Nothing)
						If tempVar Then
							Dim importedPar = dcCaptured.Import(Of Paragraph)(par, True, session)
							importedSection.Blocks.Add(importedPar)
						Else
							Dim tempVar2 As Boolean = TypeOf elementFrame.Element Is Table
							Dim table As Table = If(tempVar2, CType(elementFrame.Element, Table), Nothing)
							If tempVar2 Then
								Dim importedTable = dcCaptured.Import(Of Table)(table, True, session)
								importedSection.Blocks.Add(importedTable)
							End If
						End If
					End If
				Next elementFrame
				pageIndex += 1
			Next page
			dcCaptured.Save(outFile)
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
