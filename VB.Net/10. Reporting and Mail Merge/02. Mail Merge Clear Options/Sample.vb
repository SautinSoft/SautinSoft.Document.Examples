Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Imports SautinSoft.Document
Imports SautinSoft.Document.MailMerging

Namespace Sample
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			MailMergeWithClearOptions()
		End Sub
		''' <summary>
		''' Shows how use ClearOptions - remove specific elements if no data has been imported into them.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/mail-merge-clear-options-csharp-vb.php
		''' </remarks>
		Private Shared Sub MailMergeWithClearOptions()
			Dim document As DocumentCore = DocumentCore.Load("..\..\..\MailMergeClearOptions.docx")

			' Example 1: Remove fields for which no data has been found in the mail merge data source.
			Dim dataSource1 = New With {Key .Field2 = "Some text"}
			document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveUnusedFields
			document.MailMerge.Execute(dataSource1, "Example1")


			' Example 2: Remove paragraphs contained merge fields but none of them has been merged.
			Dim dataSource2 = New With {
				Key .Field1 = String.Empty,
				Key .Field2 = "Some text"
			}

			document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyParagraphs
			document.MailMerge.Execute(dataSource2, "Example2")


			' Example 3: Remove table rows contained merge fields but none of them has been merged.
			Dim dataSource3 = New With {
				Key .Field1 = "Some text 1",
				Key .Field2 = DirectCast(Nothing, String),
				Key .Field3 = "Some text 3"
			}

			document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyTableRows
			document.MailMerge.Execute(dataSource3, "Example3")


			' Example 4: Remove ranges into which no field has been merged.
			Dim dataSource4 = New With {
				Key .TotalRecords = 2,
				Key .Records = New Object() {
					New With {
						Key .Record = 1,
						Key .Text = "Some text 1"
					},
					New With {
						Key .Record = DirectCast(Nothing, Object),
						Key .Text = String.Empty
					},
					New With {
						Key .Record = 3,
						Key .Text = "Some text 3"
					}
				}
			}
			document.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyRanges
			document.MailMerge.Execute(dataSource4, "Example4")

			Dim resultPath As String = "ClearOptions.docx"

			' Save the output to file.
			document.Save(resultPath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
