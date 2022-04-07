Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Module Sample
    Sub Main()
        InsertTextByCoordinates()
    End Sub

	''' <summary>
	''' How to insert a text in the existing PDF, DOCX, any document by specific (x,y) coordinates
	''' </summary>
	''' <remarks>
	''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-text-in-the-existing-pdf-docx-document-by-specific-x-y-coordinates-net-csharp-vb.php
	''' </remarks>
	Sub InsertTextByCoordinates()
		' Let us say, we want to insert the text "Hello World!" into:
		' the pages 2,3;
		' 50mm - from the left;
		' 30mm - from the top.
		' Also the text must be inserted Behind the existing text.
		Dim inpFile As String = "..\..\..\example.docx"
		Dim outFile As String = "Result.docx"

		' 1. Load an existing document
		Dim dc As DocumentCore = DocumentCore.Load(inpFile)

		' 2. Get document pages
		Dim paginator = dc.GetPaginator(New PaginatorOptions() With {.UpdateFields = True})
		Dim pages = paginator.Pages

		' 3. Check that we at least 3 pages in our document.
		If pages.Count < 3 Then
			Console.WriteLine("The document contains less than 3 pages!")
			Console.ReadKey()
			Return
		End If
		' 50mm - from the left;            
		' 30mm - from the top.
		Dim posFromLeft As Single = 50.0F
		Dim posFromTop As Single = 30.0F

		' Insert the text "Hello World!" into the page 2.
		Dim text1 As New Run(dc, "Hello World!", New CharacterFormat() With {
			.Size = 36,
			.FontColor = Color.Red,
			.FontName = "Arial"
		})
		InsertShape(dc, pages(1), text1, posFromLeft, posFromTop)

		' Insert the text "Hej Världen!" into the page 3.
		Dim text2 As New Run(dc, "Hej Världen!", New CharacterFormat() With {
			.Size = 36,
			.FontColor = Color.Orange,
			.FontName = "Arial"
		})
		InsertShape(dc, pages(2), text2, posFromLeft, posFromTop)

		' 4. Save the document back.
		dc.Save(outFile)
		System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
	End Sub
	''' <summary>
	''' Inserts a Shape with text into the specific page using coordinates.
	''' </summary>
	''' <param name="dc">The document.</param>
	''' <param name="page">The specific page to insert the Shape.</param>
	''' <param name="text">The formatted text (Run object).</param>
	''' <param name="posFromLeftMm">The distance in mm from left corner of the page</param>
	''' <param name="posFromTopMm">The distance in mm from top corner of the page</param>
	Sub InsertShape(ByVal dc As DocumentCore, ByVal page As DocumentPage, ByVal text As Run, ByVal posFromLeftMm As Single, ByVal posFromTopMm As Single)
		Dim hp As New HorizontalPosition(posFromLeftMm, LengthUnit.Millimeter, HorizontalPositionAnchor.Page)
		Dim vp As New VerticalPosition(posFromTopMm, LengthUnit.Millimeter, VerticalPositionAnchor.Page)
		' 100 x 30 mm
		Dim shapeWidth As Single = 100.0F
		Dim shapeHeight As Single = 30.0F
		Dim size As SautinSoft.Document.Drawing.Size = New Size(shapeWidth, shapeHeight, LengthUnit.Millimeter)
		Dim shape As New Shape(dc, New FloatingLayout(hp, vp, size))
		shape.Text.Blocks.Add(New Paragraph(dc, text))
		' Set shape Behind the text.
		TryCast(shape.Layout, FloatingLayout).WrappingStyle = WrappingStyle.BehindText
		' Remove the shape borders.
		shape.Outline.Fill.SetEmpty()
		' Insert shape into the page.
		page.Content.End.Insert(shape.Content)
	End Sub

End Module