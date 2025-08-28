Imports System
Imports System.IO
Imports System.Collections.Generic
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports SautinSoft.Document.MailMerging
Imports Newtonsoft.Json
Imports System.Runtime.InteropServices

''' <summary>
''' Generates a report in PDF format (PDF/A) based on JSON data and .docx template.
''' </summary>
''' <remarks>
''' See details at: https://www.sautinsoft.com/products/document/help/net/developer-guide/mail-merge-generate-pdf-report-from-json-data-net-csharp-vb.php
''' </remarks>
Namespace CatBreedReportApp
    Friend Class CatBreed
        Public Property Title As String
        Public Property Description As String
        Public Property PictUrl As String
        ''' <summary>
        ''' Weight in lb. (Fields in template: WeightFrom, WeightTo). Here we are using a tuple.
        ''' </summary>
        Public Property Weight As (Integer, Integer)
    End Class
    Friend Class Program
        Public Shared Sub Main(args As String())
            ' Get your free key here:   
            ' https://sautinsoft.com/start-for-free/

            ' 1. Get json data
            Dim json As String = CreateJsonObject()

            ' 2. Show json to Console.
            Console.WriteLine(json)

            ' 3. Generate report based on .docx template and json.
            GeneratePdfReport(json)
        End Sub

        Public Shared Sub GeneratePdfReport(json As String)
            ' Get data from json.            
            Dim cats = JsonConvert.DeserializeObject(Of List(Of CatBreed))(json)

            ' Load the template document.
            Dim templatePath = "..\..\..\cats-template.docx"

            Dim dc = DocumentCore.Load(templatePath)

            ' To be able to mail merge from your own data source, it must be wrapped into an object that implements the IMailMergeDataSource interface.
            Dim customDataSource As CustomMailMergeDataSource = New CustomMailMergeDataSource(cats)

            ' Decorate each cat beed by by appropriate picture.
            ' Set picture width to 80 mm, height to Auto.
            AddHandler dc.MailMerge.FieldMerging, Sub(senderFM, eFM)
                                                      ' Insert an icon before the product name
                                                      If Equals(eFM.RangeName, "CatBreed") AndAlso Equals(eFM.FieldName, "PictUrl") Then
                                                          eFM.Inlines.Clear()
                                                          Dim pictPath As String = eFM.Value.ToString()
                                                          Dim pict As Picture = New Picture(dc, pictPath)
                                                          Dim kWH As Double = 1.0F
                                                          Dim desiredWidthMm As Double = 80

                                                          If pict.Layout.Size.Width > 0 AndAlso pict.Layout.Size.Height > 0 Then kWH = pict.Layout.Size.Width / pict.Layout.Size.Height

                                                          pict.Layout = New InlineLayout(New Size(desiredWidthMm, desiredWidthMm / kWH, LengthUnit.Millimeter))

                                                          eFM.Inlines.Add(pict)
                                                          eFM.Cancel = False
                                                      End If
                                                  End Sub

            ' Execute the mail merge.
            dc.MailMerge.Execute(customDataSource)

            Dim resultPath = "CatBreeds.pdf"

            ' Save the output to file
            Dim so As PdfSaveOptions = New PdfSaveOptions() With {
                .Compliance = PdfCompliance.PDF_A1a
            }

            dc.Save(resultPath, so)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
        End Sub
        Public Shared Function CreateJsonObject() As String
            Dim json = String.Empty
            Dim cats As List(Of CatBreed) = New List(Of CatBreed) From {
                New CatBreed() With {
                    .Title = "Australian Mist",
                    .Description = "The Australian Mist (formerly known as the Spotted Mist) is a breed of cat developed in Australia.",
                    .PictUrl = "australian-mist.jpg",
                    .Weight = (8, 15)
                },
                New CatBreed() With {
                    .Title = "Maine Coon",
                    .Description = "The Maine Coon is a large domesticated cat breed. It has a distinctive physical appearance and valuable hunting skills.",
                    .PictUrl = "maine-coon.png",
                    .Weight = (13, 18)
                },
                New CatBreed() With {
                    .Title = "Scottish Fold",
                    .Description = "The original Scottish Fold was a white barn cat named Susie, who was found at a farm near Coupar Angus in Perthshire, Scotland, in 1961.",
                    .PictUrl = "scottish-fold.jpg",
                    .Weight = (9, 13)
                },
                New CatBreed() With {
                    .Title = "Oriental Shorthair",
                    .Description = "The Oriental Shorthair is a breed of domestic cat that is developed from and closely related to the Siamese cat.",
                    .PictUrl = "oriental-shorthair.jpg",
                    .Weight = (8, 12)
                },
                New CatBreed() With {
                    .Title = "Bengal cat",
                    .Description = "The earliest mention of an Asian leopard cat × domestic cross was in 1889, when Harrison Weir wrote of them in Our Cats and ...",
                    .PictUrl = "bengal-cat.jpg",
                    .Weight = (10, 15)
                },
                New CatBreed() With {
                    .Title = "Russian Blue",
                    .Description = "The Russian Blue is a naturally occurring breed that may have originated in the port of Arkhangelsk in Russia.",
                    .PictUrl = "russian-blue.jpg",
                    .Weight = (8, 15)
                },
                New CatBreed() With {
                    .Title = "Mongrel cat",
                    .Description = "A mongrel, mutt or mixed-breed cat is a cat that does not belong to one officially recognized breed, but he's cool and gentle!",
                    .PictUrl = "mongrel-cat.jpg",
                    .Weight = (8, 16)
                }
            }

            ' Generate full path for the cat's pictures.
            Dim pictDirectory = Path.GetFullPath("..\..\..\picts\")
            For Each cb In cats
                cb.PictUrl = Path.Combine(pictDirectory, cb.PictUrl)
            Next

            ' Make serialization to JSON format.            
            json = JsonConvert.SerializeObject(cats)
            Return json
        End Function

        ''' <summary>
        ''' A custom mail merge data source that allows SautinSoft.Document to retrieve data from CatBeeds objects.
        ''' </summary>
        Public Class CustomMailMergeDataSource
            Implements IMailMergeDataSource
            Private ReadOnly _cats As List(Of CatBreed)
            Private _recordIndex As Integer

            ''' <summary>
            ''' The name of the data source. 
            ''' </summary>
            Public ReadOnly Property Name As String Implements IMailMergeDataSource.Name
                Get
                    Return "CatBreed"
                End Get
            End Property

            ''' <summary>
            ''' SautinSoft.Document calls this method to get a value for every data field.
            ''' </summary>
            Public Function TryGetValue(valueName As String, <Out> ByRef value As Object) As Boolean Implements IMailMergeDataSource.TryGetValue
                Select Case valueName
                    Case "Title"
                        value = _cats(_recordIndex).Title
                        Return True
                    Case "Description"
                        value = _cats(_recordIndex).Description
                        Return True
                    Case "PictUrl"
                        value = _cats(_recordIndex).PictUrl
                        Return True
                    Case "WeightFrom"
                        value = _cats(_recordIndex).Weight.Item1
                        Return True
                    Case "WeightTo"
                        value = _cats(_recordIndex).Weight.Item2
                        Return True
                    Case Else
                        ' A field with this name was not found
                        value = Nothing
                        Return False
                End Select
            End Function

            ''' <summary>
            ''' A standard implementation for moving to a next record in a collection.
            ''' </summary>
            Public Function MoveNext() As Boolean Implements IMailMergeDataSource.MoveNext
                Return Threading.Interlocked.Increment(_recordIndex) < _cats.Count
            End Function

            Public Function GetChildDataSource(sourceName As String) As IMailMergeDataSource Implements IMailMergeDataSource.GetChildDataSource
                Return Nothing
            End Function
            Public Sub New(cats As List(Of CatBreed))
                _cats = cats
                ' When the data source is initialized, it must be positioned before the first record.
                _recordIndex = -1
            End Sub
        End Class
    End Class
End Namespace
