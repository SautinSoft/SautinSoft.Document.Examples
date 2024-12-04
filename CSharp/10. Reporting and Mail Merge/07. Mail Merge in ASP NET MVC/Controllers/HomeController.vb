Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports Microsoft.AspNetCore.Hosting
Imports Microsoft.AspNetCore.Mvc
Imports DocumentCoreMvc.Models
Imports SautinSoft.Document

Namespace DocumentCoreMvc.Controllers
	Public Class HomeController
		Inherits Controller

		Private ReadOnly environment As IWebHostEnvironment

		Public Sub New(ByVal environment As IWebHostEnvironment)
			Me.environment = environment
		End Sub

		Public Function Index() As IActionResult
			Return View(New InvoiceModel())
		End Function

		Public Function Download(ByVal model As InvoiceModel) As FileStreamResult
			' Load template document.
			Dim path As System.String = System.IO.Path.Combine(Me.environment.ContentRootPath, "InvoiceWithFields.docx")
			Dim document = DocumentCore.Load(path)

			' Execute mail merge process.
			document.MailMerge.Execute(model)

			' Save document in specified file format.
			Dim stream = New MemoryStream()
			document.Save(stream, model.Options)

			' Set the stream position to the beginning of the file.
			'fileStream.Seek(0, SeekOrigin.Begin);

			stream.Seek(0, 0)
			' Download file.
			Return File(stream, model.Options.ContentType, $"OutputFromView.{model.Format.ToLower()}")

		End Function

	   ' [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
		Public Function [Error]() As IActionResult
			Return View(New ErrorViewModel() With {.RequestId = If(Activity.Current?.Id, HttpContext.TraceIdentifier)})
		End Function
	End Class
End Namespace

Namespace DocumentCoreMvc.Models
	Public Class InvoiceModel
		Public Property Number As Integer = 1
		Public Property [Date] As DateTime = DateTime.Today
		Public Property Company As String = "Springfield Nuclear Power Plant (in Sector 7-G)"
		Public Property Address As String = "742 Evergreen Terrace, Springfield, United States"
		Public Property Name As String = "Homer Simpson"
		Public Property Format As String = "DOCX"
		Public ReadOnly Property Options As SaveOptions
			Get
				Return Me.FormatMappingDictionary(Me.Format)
			End Get
		End Property
		Public ReadOnly Property FormatMappingDictionary As IDictionary(Of String, SaveOptions)
			Get
				Return New Dictionary(Of String, SaveOptions)() From {
					{"DOCX", New DocxSaveOptions()},
					{
						"HTML", New HtmlFixedSaveOptions With {.EmbedImages =True}
					},
					{"RTF", New RtfSaveOptions()},
					{"TXT", New TxtSaveOptions()},
					{"PDF", New PdfSaveOptions()}
				}
			End Get
		End Property
	End Class
End Namespace
