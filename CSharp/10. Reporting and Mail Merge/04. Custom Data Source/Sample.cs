using System.Collections.Generic;
using SautinSoft.Document;
using SautinSoft.Document.MailMerging;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            CustomDataSource();
        }
        /// <summary>
        /// Generate reports using a custom data source (collection of custom classes Actor and Order).
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/mail-merge-custom-data-source-net-csharp-vb.php
        /// </remarks>
        static void CustomDataSource()
        {
            // Populate some data that we will use in the mail merge.
            List<Actor> actors = new List<Actor>();
            actors.Add(new Actor("Arnold Schwarzenegger", "12989 Chalon Road, Los Angeles, CA 90049"));
            actors.Add(new Actor("Sylvester Stallone", "30 Beverly Park Terrace, Beverly Hills, CA 90210"));

            // Populate some data for nesting in the mail merge.
            actors[0].Orders.Add(new Order("Bowflex SelectTech 1090 Adjustable Dumbbell", 2));
            actors[0].Orders.Add(new Order("Gold's Gym Kettlebell Kit, 5-15 Lbs.", 1));
            actors[1].Orders.Add(new Order("Weider Cast Iron Olympic Hammertone Weight Set, 300 Lb.", 1));

            // Load the template document.
            DocumentCore dc = DocumentCore.Load(@"..\..\..\OrdersTemplate.docx");

            // To be able to mail merge from your own data source, it must be wrapped into an object that implements the IMailMergeDataSource interface.
            CustomMailMergeDataSource customDataSource = new CustomMailMergeDataSource(actors);

            // Execute the mail merge.
            dc.MailMerge.Execute(customDataSource);

            string resultPath = "Orders.docx";

            // Save the output to file.
            dc.Save(resultPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }

        /// <summary>
        /// An example of a class that contain actor's data.
        /// </summary>
        public class Actor
        {
            private string _fullName;
            private string _address;
            private List<Order> _orders;

            public string FullName
            {
                get { return _fullName; }
                set { _fullName = value; }
            }

            public string Address
            {
                get { return _address; }
                set { _address = value; }
            }

            public List<Order> Orders
            {
                get { return _orders; }
                set { _orders = value; }
            }

            public Actor(string fullName, string address)
            {
                _fullName = fullName;
                _address = address;
                _orders = new List<Order>();
            }

        }         

        /// <summary>
        /// An example of a class that contain order's data.
        /// </summary>
        public class Order
        {
            private string _name;
            private int _quantity;

            public string Name
            {
                get { return _name; }
                set { _name = value; }
            }

            public int Quantity
            {
                get { return _quantity; }
                set { _quantity = value; }
            }

            public Order(string name, int quantity)
            {
                _name = name;
                _quantity = quantity;
            }
        }                     

        /// <summary>
        /// A custom mail merge data source that allows SautinSoft.Document to retrieve data from Actor objects.
        /// </summary>
        public class CustomMailMergeDataSource : IMailMergeDataSource
        {
            private readonly List<Actor> _actors;
            private int _recordIndex;

            /// <summary>
            /// The name of the data source. 
            /// </summary>
            public string Name
            {
                get { return "Actor"; }
            }

            /// <summary>
            /// SautinSoft.Document calls this method to get a value for every data field.
            /// </summary>
            public bool TryGetValue(string valueName, out object value)
            {
                switch (valueName)
                {
                    case "FullName":
                        value = _actors[_recordIndex].FullName;
                        return true;
                    case "Address":
                        value = _actors[_recordIndex].Address;
                        return true;
                    default:
                        // A field with this name was not found
                        value = null;
                        return false;
                }
            }

            /// <summary>
            /// A standard implementation for moving to a next record in a collection.
            /// </summary>
            public bool MoveNext()
            {
                return (++_recordIndex < _actors.Count);
            }

            public IMailMergeDataSource GetChildDataSource(string sourceName)
            {
                switch (sourceName)
                {
                    case "Order":
                        return new OrderMailMergeDataSource(_actors[_recordIndex].Orders);
                    default:
                        return null;
                }
            }

            public CustomMailMergeDataSource(List<Actor> actors)
            {
                _actors = actors;
                // When the data source is initialized, it must be positioned before the first record.
                _recordIndex = -1;
            }
        }

        public class OrderMailMergeDataSource : IMailMergeDataSource
        {
            private readonly List<Order> _orders;
            private int _recordIndex;

            /// <summary>
            /// The name of the data source. 
            /// </summary>
            public string Name
            {
                get { return "Order"; }
            }

            /// <summary>
            /// SautinSoft.Document calls this method to get a value for every data field.
            /// </summary>
            public bool TryGetValue(string valueName, out object value)
            {
                switch (valueName)
                {
                    case "Name":
                        value = _orders[_recordIndex].Name;
                        return true;
                    case "Quantity":
                        value = _orders[_recordIndex].Quantity;
                        return true;
                    default:
                        // A field with this name was not found
                        value = null;
                        return false;
                }
            }

            /// <summary>
            /// A standard implementation for moving to a next record in a collection.
            /// </summary>
            public bool MoveNext()
            {
                return (++_recordIndex < _orders.Count);
            }

            // Return null because Order haven't any child elements.
            public IMailMergeDataSource GetChildDataSource(string tableName)
            {
                return null;
            }

            public OrderMailMergeDataSource(List<Order> orders)
            {
                _orders = orders;
                // When the data source is initialized, it must be positioned before the first record.
                _recordIndex = -1;
            }
        }
    }
}