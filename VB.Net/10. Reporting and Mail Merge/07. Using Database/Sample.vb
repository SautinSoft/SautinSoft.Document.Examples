Imports System
Imports System.Data
Imports System.IO
Imports SautinSoft.Document
Imports System.Linq
Imports System.Globalization
Imports System.Data.OleDb

Namespace Sample
	Friend Class Sample
		Shared Sub Main(ByVal args() As String)
			MailMergeUsingDatabase()
		End Sub
		''' <summary>
		''' How to generate a Report (Mail Merge) using a DOCX template and Database as a data source.
		''' </summary>
		''' <remarks>
		''' See details at: https://www.sautinsoft.com/products/document/help/net/developer-guide/mail-merge-using-database-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub MailMergeUsingDatabase()
			' Load the DOCX-template document. 
			Dim document As DocumentCore = DocumentCore.Load("..\..\..\OrdersDbTemplate.docx")

			' Create a data source.
			Dim dataSet As New DataSet()

			' Execute query for retrieving data of Orders table.
			Dim orders As DataTable = ExecuteSQL(String.Join(" ", "SELECT DISTINCT Orders.ShipName, Orders.ShipAddress, Orders.ShipCity,", "Orders.ShipRegion, Orders.ShipPostalCode, Orders.ShipCountry, Orders.CustomerID,", "Customers.CompanyName AS Customers_CompanyName, Customers.Address, Customers.City,", "Customers.Region, Customers.PostalCode, Customers.Country,", "[FirstName] & "" "" & [LastName] AS SalesPerson, Orders.OrderID, Orders.OrderDate,", "Orders.RequiredDate, Orders.ShippedDate, Shippers.CompanyName AS Shippers_CompanyName", "FROM Shippers INNER JOIN (Employees INNER JOIN (Customers INNER JOIN Orders ON", "Customers.CustomerID = Orders.CustomerID) ON Employees.EmployeeID = Orders.EmployeeID)", "ON Shippers.ShipperID = Orders.ShipVia"))
			orders.TableName = "Orders"
			dataSet.Tables.Add(orders)

			' Execute query for retrieving data of OrderDetails table.
			Dim orderDetails As DataTable = ExecuteSQL(String.Join(" ", "SELECT DISTINCTROW Orders.OrderID, [Order Details].ProductID, Products.ProductName,", "[Order Details].UnitPrice, [Order Details].Quantity, [Order Details].Discount,", "([Order Details].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100 AS TotalPrice", "FROM Products INNER JOIN (Orders INNER JOIN [Order Details] ON", "Orders.OrderID = [Order Details].OrderID) ON Products.ProductID = [Order Details].ProductID"))
			orderDetails.TableName = "OrderDetails"
			dataSet.Tables.Add(orderDetails)

			' Add parent-child relation.
			orders.ChildRelations.Add("OrderDetails", orders.Columns("OrderID"), orderDetails.Columns("OrderID"))

			' Calculate and fill Total.
			document.MailMerge.Execute(New With {Key .Total = (CDbl(orderDetails.Rows.Cast(Of DataRow)().Sum(Function(item) CDbl(item("TotalPrice")))))})

			' Calculate Subtotal for the each order.
			AddHandler document.MailMerge.FieldMerging, Sub(sender, e)
															If e.RangeName = "Order" AndAlso e.FieldName = "Subtotal" Then
																e.Inline = New Run(e.Document, CDbl(dataSet.Tables("Orders").Rows(e.RecordNumber - 1).GetChildRows("OrderDetails").Sum(Function(item) CDbl(item("TotalPrice")))).ToString("$#,##0.00", CultureInfo.InvariantCulture))
																e.Cancel = False
															End If
														End Sub

			' Execute the Mail Merge.
			' Note: As the name of the region in the template (Order) is different from the name of the table (Orders), we explicitly specify the name of the region.
			document.MailMerge.Execute(orders, "Order")

			Dim resultPath As String = "Orders.docx"

			' Save the output to file.
			document.Save(resultPath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub

		''' <summary>
		''' Utility function that creates a connection, executes the sql-query and 
		''' return the result in a DataTable.
		''' </summary>
		Private Shared Function ExecuteSQL(ByVal sqlText As String) As DataTable
			' Open the database connection.
			Dim str As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\..\..\Northwind.mdb"
			Dim conn As New OleDbConnection(str)
			conn.Open()

			' Create and execute a command.
			Dim cmd As New OleDbCommand(sqlText, conn)
			Dim da As New OleDbDataAdapter(cmd)
			Dim table As New DataTable()
			da.Fill(table)

			' Close the database.
			conn.Close()

			Return table
		End Function
	End Class
End Namespace
