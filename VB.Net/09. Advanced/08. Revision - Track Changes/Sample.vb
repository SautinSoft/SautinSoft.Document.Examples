Imports SautinSoft.Document

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			Revision()
		End Sub
		''' <summary>
		''' Shows how to work with revisions.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/revision-track-changes-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub Revision()
			Dim dc As DocumentCore = DocumentCore.Load("..\..\..\example.docx")

			' Accepting the deletion revision will assimilate it into the paragraph's inlines and remove them from the collection.
			dc.Revisions(0).Accept()

			' The second insertion revision is now at index 0, which we can reject to ignore and discard it.
			dc.Revisions(0).Reject()

			' Now we have two revisions in the list items, we accept them all.
			dc.Revisions.AcceptAll()

			dc.Save("result.pdf")
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo("result.pdf") With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
