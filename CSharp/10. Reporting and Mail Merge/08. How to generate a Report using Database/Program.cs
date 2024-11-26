using System.Globalization;
using System.Data;
using System.Data.OleDb;
using SautinSoft.Document;
using System.Linq;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            MailMergeUsingDatabase();
        }
        /// <summary>
        /// How to generate a Report (Mail Merge) using a DOCX template and Database as a data source.
        /// </summary>
        /// <remarks>
        /// See details at: https://www.sautinsoft.com/products/document/help/net/developer-guide/mail-merge-using-database-net-csharp-vb.php
        /// </remarks>
        static void MailMergeUsingDatabase()
        {
            // Load the DOCX-template document. 
            DocumentCore document = DocumentCore.Load(@"..\..\..\OrdersDbTemplate.docx");

            // Create a data source.
            DataSet dataSet = new DataSet();

            // Execute query for retrieving data of Orders table.
            DataTable orders = ExecuteSQL(string.Join(" ",
                "SELECT DISTINCT Orders.ShipName, Orders.ShipAddress, Orders.ShipCity,",
                "Orders.ShipRegion, Orders.ShipPostalCode, Orders.ShipCountry, Orders.CustomerID,",
                "Customers.CompanyName AS Customers_CompanyName, Customers.Address, Customers.City,",
                "Customers.Region, Customers.PostalCode, Customers.Country,",
                @"[FirstName] & "" "" & [LastName] AS SalesPerson, Orders.OrderID, Orders.OrderDate,",
                "Orders.RequiredDate, Orders.ShippedDate, Shippers.CompanyName AS Shippers_CompanyName",
                "FROM Shippers INNER JOIN (Employees INNER JOIN (Customers INNER JOIN Orders ON",
                "Customers.CustomerID = Orders.CustomerID) ON Employees.EmployeeID = Orders.EmployeeID)",
                "ON Shippers.ShipperID = Orders.ShipVia"));
            orders.TableName = "Orders";
            dataSet.Tables.Add(orders);

            // Execute query for retrieving data of OrderDetails table.
            DataTable orderDetails = ExecuteSQL(string.Join(" ",
                "SELECT DISTINCTROW Orders.OrderID, [Order Details].ProductID, Products.ProductName,",
                "[Order Details].UnitPrice, [Order Details].Quantity, [Order Details].Discount,",
                "([Order Details].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100 AS TotalPrice",
                "FROM Products INNER JOIN (Orders INNER JOIN [Order Details] ON",
                "Orders.OrderID = [Order Details].OrderID) ON Products.ProductID = [Order Details].ProductID"));
            orderDetails.TableName = "OrderDetails";
            dataSet.Tables.Add(orderDetails);

            // Add parent-child relation.
            orders.ChildRelations.Add("OrderDetails", orders.Columns["OrderID"], orderDetails.Columns["OrderID"]);

            // Calculate and fill Total.
            document.MailMerge.Execute(
                new
                {
                    Total = ((double)orderDetails.Rows.Cast<DataRow>().Sum(item => (double)item["TotalPrice"])),
                });

            // Calculate Subtotal for the each order.
            document.MailMerge.FieldMerging += (sender, e) =>
            {
                if (e.RangeName == "Order" && e.FieldName == "Subtotal")
                {
                    e.Inline = new Run(e.Document, ((double)dataSet.Tables["Orders"].Rows[e.RecordNumber - 1].
                        GetChildRows("OrderDetails").Sum(item => (double)item["TotalPrice"])).
                        ToString("$#,##0.00", CultureInfo.InvariantCulture));
                    e.Cancel = false;
                }
            };

            // Execute the Mail Merge.
            // Note: As the name of the region in the template (Order) is different from the name of the table (Orders), we explicitly specify the name of the region.
            document.MailMerge.Execute(orders, "Order");

            string resultPath = @"..\..\..\Orders.docx";

            // Save the output to file.
            document.Save(resultPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }

        /// <summary>
        /// Utility function that creates a connection, executes the sql-query and 
        /// return the result in a DataTable.
        /// </summary>
        static DataTable ExecuteSQL(string sqlText)
        {
            // Open the database connection.
            string str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + @"..\..\..\Northwind.mdb";
            OleDbConnection conn = new OleDbConnection(str);
            conn.Open();

            // Create and execute a command.
            OleDbCommand cmd = new OleDbCommand(sqlText, conn);
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataTable table = new DataTable();
            da.Fill(table);

            // Close the database.
            conn.Close();

            return table;
        }
    }
}