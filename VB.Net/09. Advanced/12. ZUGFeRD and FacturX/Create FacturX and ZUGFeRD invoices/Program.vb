Imports System.Globalization
Imports System.Data
Imports System.Data.OleDb
Imports SautinSoft.Document
Imports System.Linq
Imports System.IO

Namespace Sample
    Class Sample
        Shared Sub Main(args As String())
            Create_Factur_ZUGFeRD_invoices()
        End Sub

        Private Shared Sub Create_Factur_ZUGFeRD_invoices()
            Dim document As DocumentCore = DocumentCore.Load("..\..\..\OrdersDbTemplate.docx")
            Dim xmlInfo As String = File.ReadAllText("..\..\..\info.xml")
            Dim resultPath As String = "..\..\..\Orders.pdf"
            Dim dataSet As DataSet = New DataSet()
            Dim orders As DataTable = ExecuteSQL(String.Join(" ", "SELECT DISTINCT Orders.ShipName, Orders.ShipAddress, Orders.ShipCity,", "Orders.ShipRegion, Orders.ShipPostalCode, Orders.ShipCountry, Orders.CustomerID,", "Customers.CompanyName AS Customers_CompanyName, Customers.Address, Customers.City,", "Customers.Region, Customers.PostalCode, Customers.Country,", "[FirstName] & "" "" & [LastName] AS SalesPerson, Orders.OrderID, Orders.OrderDate,", "Orders.RequiredDate, Orders.ShippedDate, Shippers.CompanyName AS Shippers_CompanyName", "FROM Shippers INNER JOIN (Employees INNER JOIN (Customers INNER JOIN Orders ON", "Customers.CustomerID = Orders.CustomerID) ON Employees.EmployeeID = Orders.EmployeeID)", "ON Shippers.ShipperID = Orders.ShipVia"))
            orders.TableName = "Orders"
            dataSet.Tables.Add(orders)
            Dim orderDetails As DataTable = ExecuteSQL(String.Join(" "c,
                "SELECT DISTINCTROW Orders.OrderID, [Order Details].ProductID, Products.ProductName,",
                "[Order Details].UnitPrice, [Order Details].Quantity, [Order Details].Discount,",
                "([Order Details].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100 AS TotalPrice",
                "FROM Products INNER JOIN (Orders INNER JOIN [Order Details] ON",
                "Orders.OrderID = [Order Details].OrderID) ON Products.ProductID = [Order Details].ProductID"))
            orderDetails.TableName = "OrderDetails"
            dataSet.Tables.Add(orderDetails)

            ' Add parent-child relation.
            orders.ChildRelations.Add("OrderDetails", orders.Columns("OrderID"), orderDetails.Columns("OrderID"))

            document.MailMerge.Execute(
                New With {
                    .Total = CDbl(orderDetails.Rows.Cast(Of DataRow)().Sum(Function(item) CDbl(item("TotalPrice"))))
                })


            document.MailMerge.Execute(orders, "Order")
            Dim pdfSO As PdfSaveOptions = New PdfSaveOptions() With {
                .FacturXXML = xmlInfo
            }
            document.Save(resultPath, pdfSO)
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {
                .UseShellExecute = True
            })
        End Sub

        Private Shared Function ExecuteSQL(ByVal sqlText As String) As DataTable
            Dim str As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & "..\..\..\Northwind.mdb"
            Dim conn As OleDbConnection = New OleDbConnection(str)
            conn.Open()
            Dim cmd As OleDbCommand = New OleDbCommand(sqlText, conn)
            Dim da As OleDbDataAdapter = New OleDbDataAdapter(cmd)
            Dim table As DataTable = New DataTable()
            da.Fill(table)
            conn.Close()
            Return table
        End Function
    End Class
End Namespace
