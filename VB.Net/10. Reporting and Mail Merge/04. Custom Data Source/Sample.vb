Imports System.Collections.Generic
Imports SautinSoft.Document
Imports SautinSoft.Document.MailMerging

Namespace Sample
    Friend Class Sample
        Shared Sub Main(ByVal args() As String)
            CustomDataSource()
        End Sub
        ''' <summary>
        ''' Generate reports using a custom data source (collection of custom classes Actor and Order).
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/mail-merge-custom-data-source-net-csharp-vb.php
        ''' </remarks>
        Private Shared Sub CustomDataSource()
            ' Populate some data that we will use in the mail merge.
            Dim actors As New List(Of Actor)()
            actors.Add(New Actor("Arnold Schwarzenegger", "12989 Chalon Road, Los Angeles, CA 90049"))
            actors.Add(New Actor("Sylvester Stallone", "30 Beverly Park Terrace, Beverly Hills, CA 90210"))

            ' Populate some data for nesting in the mail merge.
            actors(0).Orders.Add(New Order("Bowflex SelectTech 1090 Adjustable Dumbbell", 2))
            actors(0).Orders.Add(New Order("Gold's Gym Kettlebell Kit, 5-15 Lbs.", 1))
            actors(1).Orders.Add(New Order("Weider Cast Iron Olympic Hammertone Weight Set, 300 Lb.", 1))

            ' Load the template document.
            Dim dc As DocumentCore = DocumentCore.Load("..\..\..\OrdersTemplate.docx")

            ' To be able to mail merge from your own data source, it must be wrapped into an object that implements the IMailMergeDataSource interface.
            'INSTANT VB NOTE: The variable customDataSource was renamed since Visual Basic does not handle local variables named the same as class members well:
            Dim customDataSource_Renamed As New CustomMailMergeDataSource(actors)

            ' Execute the mail merge.
            dc.MailMerge.Execute(customDataSource_Renamed)

            Dim resultPath As String = "Orders.docx"

            ' Save the output to file.
            dc.Save(resultPath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
        End Sub

        ''' <summary>
        ''' An example of a class that contain actor's data.
        ''' </summary>
        Public Class Actor
            Private _fullName As String
            Private _address As String
            Private _orders As List(Of Order)

            Public Property FullName() As String
                Get
                    Return _fullName
                End Get
                Set(ByVal value As String)
                    _fullName = value
                End Set
            End Property

            Public Property Address() As String
                Get
                    Return _address
                End Get
                Set(ByVal value As String)
                    _address = value
                End Set
            End Property

            Public Property Orders() As List(Of Order)
                Get
                    Return _orders
                End Get
                Set(ByVal value As List(Of Order))
                    _orders = value
                End Set
            End Property

            Public Sub New(ByVal fullName As String, ByVal address As String)
                _fullName = fullName
                _address = address
                _orders = New List(Of Order)()
            End Sub

        End Class

        ''' <summary>
        ''' An example of a class that contain order's data.
        ''' </summary>
        Public Class Order
            Private _name As String
            Private _quantity As Integer

            Public Property Name() As String
                Get
                    Return _name
                End Get
                Set(ByVal value As String)
                    _name = value
                End Set
            End Property

            Public Property Quantity() As Integer
                Get
                    Return _quantity
                End Get
                Set(ByVal value As Integer)
                    _quantity = value
                End Set
            End Property

            Public Sub New(ByVal name As String, ByVal quantity As Integer)
                _name = name
                _quantity = quantity
            End Sub
        End Class

        ''' <summary>
        ''' A custom mail merge data source that allows SautinSoft.Document to retrieve data from Actor objects.
        ''' </summary>
        Public Class CustomMailMergeDataSource
            Implements IMailMergeDataSource

            Private ReadOnly _actors As List(Of Actor)
            Private _recordIndex As Integer

            ''' <summary>
            ''' The name of the data source. 
            ''' </summary>
            Public ReadOnly Property Name() As String Implements IMailMergeDataSource.Name
                Get
                    Return "Actor"
                End Get
            End Property

            ''' <summary>
            ''' SautinSoft.Document calls this method to get a value for every data field.
            ''' </summary>
            Public Function TryGetValue(ByVal valueName As String, <System.Runtime.InteropServices.Out()> ByRef value As Object) As Boolean Implements IMailMergeDataSource.TryGetValue
                Select Case valueName
                    Case "FullName"
                        value = _actors(_recordIndex).FullName
                        Return True
                    Case "Address"
                        value = _actors(_recordIndex).Address
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
                _recordIndex += 1
                'INSTANT VB WARNING: An assignment within expression was extracted from the following statement:
                'ORIGINAL LINE: return (++_recordIndex < _actors.Count);
                Return (_recordIndex < _actors.Count)
            End Function

            Public Function GetChildDataSource(ByVal sourceName As String) As IMailMergeDataSource Implements IMailMergeDataSource.GetChildDataSource
                Select Case sourceName
                    Case "Order"
                        Return New OrderMailMergeDataSource(_actors(_recordIndex).Orders)
                    Case Else
                        Return Nothing
                End Select
            End Function

            Public Sub New(ByVal actors As List(Of Actor))
                _actors = actors
                ' When the data source is initialized, it must be positioned before the first record.
                _recordIndex = -1
            End Sub
        End Class

        Public Class OrderMailMergeDataSource
            Implements IMailMergeDataSource

            Private ReadOnly _orders As List(Of Order)
            Private _recordIndex As Integer

            ''' <summary>
            ''' The name of the data source. 
            ''' </summary>
            Public ReadOnly Property Name() As String Implements IMailMergeDataSource.Name
                Get
                    Return "Order"
                End Get
            End Property

            ''' <summary>
            ''' SautinSoft.Document calls this method to get a value for every data field.
            ''' </summary>
            Public Function TryGetValue(ByVal valueName As String, <System.Runtime.InteropServices.Out()> ByRef value As Object) As Boolean Implements IMailMergeDataSource.TryGetValue
                Select Case valueName
                    Case "Name"
                        value = _orders(_recordIndex).Name
                        Return True
                    Case "Quantity"
                        value = _orders(_recordIndex).Quantity
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
                _recordIndex += 1
                'INSTANT VB WARNING: An assignment within expression was extracted from the following statement:
                'ORIGINAL LINE: return (++_recordIndex < _orders.Count);
                Return (_recordIndex < _orders.Count)
            End Function

            ' Return null because Order haven't any child elements.
            Public Function GetChildDataSource(ByVal tableName As String) As IMailMergeDataSource Implements IMailMergeDataSource.GetChildDataSource
                Return Nothing
            End Function

            Public Sub New(ByVal orders As List(Of Order))
                _orders = orders
                ' When the data source is initialized, it must be positioned before the first record.
                _recordIndex = -1
            End Sub
        End Class
    End Class
End Namespace