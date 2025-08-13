Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SautinSoft.Document
Imports System
Imports System.Collections.Generic

Namespace Sample
    Friend Class Sample

        Public Shared Sub Main()
            ' Get your free trial key here:   
            ' https://sautinsoft.com/start-for-free/

            ''' <summary>
            ''' Filling out a document from a database.
            ''' </summary>
            ''' <remarks>
            ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/read_content_from_db.php
            ''' </remarks>

            Dim db_content = "
                {
                    'SomeField': 'SautinSoft',
                    'Field': 'RATES',
                    'Table1': [
                        {
                            'Name': 'John S. Director',
                            'Description': '$475.00',
                            'Bla': 'In case of deviations from the schedule, analyzes the reasons for the deviations and takes measures to correct the situation.'
                        },
                        {
                            'Name': 'Donna Di Group',
                            'Description': '$150.67',
                            'Bla': 'Plans to perform the work using own resources or the need to involve a subcontractor in order to determine the cost of designing the facility and ensure the release of documentation within the timeframe specified in the technical specifications.'
                        },
                        {
                            'Name': 'Frank H.Rogel',
                            'Description': '$623.09',
                            'Bla': 'Forms a commercial proposal for his sections and submits it to the marketing and tender department for his section of work.'
                        }
                    ],
                    'Table2': [
                        {
                            'Name': 'Name 4',
                            'Description': 'Description 4'
                        },
                        {
                            'Name': 'Name 5',
                            'Description': 'Description 5'
                        },
                        {
                            'Name': 'Name 6',
                            'Description': 'Description 6'
                        }
                    ]
                }
            "

            Dim jsonSettings As New JsonSerializerSettings()
            jsonSettings.Converters.Add(New JsonObjectConverter())

            Dim jsonSource = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(db_content, jsonSettings)

            Dim document = DocumentCore.Load("..\..\..\Template.docx")
            document.MailMerge.Execute(jsonSource)
            document.Save("..\..\..\Result.pdf")
        End Sub

        Public NotInheritable Class JsonObjectConverter
            Inherits JsonConverter
            Public Overrides Function CanConvert(objectType As Type) As Boolean
                Return objectType Is GetType(Object)
            End Function

            Public Overrides Function ReadJson(reader As JsonReader, objectType As Type, existingValue As Object, serializer As JsonSerializer) As Object
                Select Case reader.TokenType
                    Case JsonToken.StartObject
                        Return serializer.Deserialize(Of Dictionary(Of String, Object))(reader)

                    Case JsonToken.StartArray
                        Dim array = JArray.Load(reader)
                        Dim list = New List(Of Dictionary(Of String, Object))()

                        For Each item In array
                            If item.Type = JTokenType.Object Then list.Add(item.ToObject(Of Dictionary(Of String, Object))())
                        Next

                        Return list

                    Case JsonToken.String
                        Return reader.Value?.ToString()

                    Case JsonToken.Integer, JsonToken.Float
                        Return Convert.ToDouble(reader.Value)

                    Case JsonToken.Boolean
                        Return Convert.ToBoolean(reader.Value)
                    Case Else
                        Dim token = JToken.ReadFrom(reader)
                        Return token.ToObject(Of Object)()
                End Select
            End Function

            Public Overrides Sub WriteJson(writer As JsonWriter, value As Object, serializer As JsonSerializer)
                serializer.Serialize(writer, value)
            End Sub
        End Class
    End Class
End Namespace
