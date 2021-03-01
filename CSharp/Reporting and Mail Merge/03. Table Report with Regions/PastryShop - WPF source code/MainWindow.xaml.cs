using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Serialization;
using System.IO;
using System.Collections.ObjectModel;
using System.ComponentModel;
//using SautinSoft.Document;
using System.Data;
//using SautinSoft.Document.MailMerging;


namespace PastryShop
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {

        public ObservableCollection<Order> Orders { get; set; }
        public Dictionary<string, double> PastryItems { get; set; }

        public List<Person> Persons { get; set; }

        public ObservableCollection<Product> Products
        {
            get
            {
                return Orders.Last().Products;
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public Order Order
        {
            get
            {
                if (Orders != null && Orders.Count > 0)
                    return Orders[Orders.Count - 1];
                else
                    return null;
            }
            set
            {
                if (Orders != null && Orders.Count > 0)
                {
                    Orders[Orders.Count - 1] = value;
                }
                OnPropertyChanged("Order");
            }
        }
        public Person GetPerson(string name)
        {
            Person person = null;
            if (Persons.Count > 0)
            {
                person = Persons.First(p => p.Name == name);
            }
            return person;
        }

        public string ProductsHeader
        {
            get { return (string)GetValue(ProductsHeaderProperty); }
            set { SetValue(ProductsHeaderProperty, value); }
        }
        public static readonly DependencyProperty ProductsHeaderProperty =
            DependencyProperty.Register("ProductsHeader", typeof(string), typeof(MainWindow), new PropertyMetadata(null));

        public string OrderNumber
        {
            get { return (string)GetValue(OrderNumberProperty); }
            set { SetValue(OrderNumberProperty, value); }
        }
        public static readonly DependencyProperty OrderNumberProperty =
            DependencyProperty.Register("OrderNumber", typeof(string), typeof(MainWindow), new PropertyMetadata(null));


        public MainWindow()
        {
            InitializeComponent();

            // Create list with orders.
            Orders = new ObservableCollection<Order>();

            PastryItems = new Dictionary<string, double>();

            PastryItems.Add("Cherry/apple/pie", 10.0);
            PastryItems.Add("Dark/milk/white chocolate", 3.0);
            PastryItems.Add("Rice/lemon/vanilla pudding", 8.0);
            PastryItems.Add("Chocolate/strawberry/vanilla ice cream", 4.0);
            PastryItems.Add("Doughnut", 3.0);
            PastryItems.Add("Scone", 5.0);
            PastryItems.Add("White bread", 3.0);
            PastryItems.Add("Rye bread", 3.0);
            PastryItems.Add("French bread (baguette)", 4.0);
            PastryItems.Add("Hamburger bun", 4.0);
            PastryItems.Add("Hot dog bun", 2.0);
            PastryItems.Add("Roll", 2.0);
            PastryItems.Add("Croissant", 4.0);
            PastryItems.Add("Swissroll", 8.0);
            PastryItems.Add("Pancake", 1.0);
            PastryItems.Add("Spice-cake, honey-cake", 7.0);
            PastryItems.Add("Chocolate cake", 18.0);

            FillPersonsByDefault();
            FillOrdersByDefault();

            OrderNumber = String.Format("Order {0:D3}", Order.Number);
            ProductsHeader = String.Format("Products list - {0}", OrderNumber);

            // File combobox "By the Time"
            for (int h = 0; h <= 24; h++)
            {
                cmbTime.Items.Add(String.Format("{0:D2}:00", h));
                if (h != 24)
                    cmbTime.Items.Add(String.Format("{0:D2}:30", h));
            }

        }

        public void FillPersonsByDefault()
        {
            // Create fake Persons
            // https://fakeaddressgenerator.com
            Persons = new List<Person>();
            Persons.Add(new Person()
            {
                Name = "Spencer F Clements",
                Address = "17 Chapel Lane",
                City = "ARGOED",
                Phone = "077 7775 4600"
            });

            Persons.Add(new Person()
            {
                Name = "Naomi H Humphreys",
                Address = "26 Merthyr Road",
                City = "BURGHFIELD",
                Phone = "077 3399 6360"
            });

            Persons.Add(new Person()
            {
                Name = "Michael M Preston",
                Address = "27 Hendford Hill",
                City = "MOSELEY",
                Phone = "070 2532 1566"
            });

            Persons.Add(new Person()
            {
                Name = "Olivia N Dixon",
                Address = "8 Main St",
                City = "ACHNADRISH",
                Phone = "078 1863 2697"
            });


            Persons.Add(new Person()
            {
                Name = "Kieran P Sheppard",
                Address = "64 Guild Street",
                City = "LONDON",
                Phone = "078 2241 7964"
            });

            Persons.Add(new Person()
            {
                Name = "Zachary F Chambers",
                Address = "40 Shannon Way",
                City = "CHILTON STREET",
                Phone = "077 2240 7492"
            });

            Persons.Add(new Person()
            {
                Name = "Jamie E Duffy",
                Address = "99 Ballifeary Road",
                City = "BALMULLO",
                Phone = "077 1822 1757"
            });

            Persons.Add(new Person()
            {
                Name = "Connor F Atkinson",
                Address = "76 Stroude Road",
                City = "SKELBERRY",
                Phone = "070 6466 9938"
            });

        }

        public void FillOrdersByDefault()
        {
            Order o = new Order();
            o.Number = 1;
            o.Recipient = Persons[0].Name;
            o.Address = Persons[0].Address;
            o.City = Persons[0].City;
            o.Phone = Persons[0].Phone;
            o.Date = DateTime.Now.ToShortDateString();
            o.Time = "10:00";

            Random rnd = new Random();

            Product p1 = new Product()
            {
                Name = PastryItems.ElementAt(0).Key,
                Price = PastryItems.ElementAt(0).Value,
                Quantity = rnd.Next(1, 10)
            };
            p1.Recalculate();

            Product p2 = new Product()
            {
                Name = PastryItems.ElementAt(1).Key,
                Price = PastryItems.ElementAt(1).Value,
                Quantity = rnd.Next(1, 10)
            };
            p2.Recalculate();

            Product p3 = new Product()
            {
                Name = PastryItems.ElementAt(2).Key,
                Price = PastryItems.ElementAt(2).Value,
                Quantity = rnd.Next(1, 10)
            };
            p3.Recalculate();

            o.Products.Add(p1);
            o.Products.Add(p2);
            o.Products.Add(p3);

            Orders.Add(o);
        }

        private void btnAddProduct_Click(object sender, RoutedEventArgs e)
        {
            Product product = new Product();
            product.Quantity = 1;
            Products.Add(product);
        }

        private void ProductsSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox combo = sender as ComboBox;
            if (combo.IsDropDownOpen)
            {
                string selectedProductName = combo.SelectedValue.ToString();

                Product p = gridProducts.SelectedItem as Product;
                if (p == null)
                {
                    p = new Product();
                    Products.Add(p);
                }

                p.Name = selectedProductName;
                p.Price = PastryItems[p.Name];
                p.Quantity = 1;

                p.Recalculate();
                RefreshGrid();
            }
        }
        private void btnRemoveProduct_Click(object sender, RoutedEventArgs e)
        {
            if (Products.Count > 0)
            {
                // Get selected product(s)
                var pList = gridProducts.SelectedItems;

                // If we didn't select any product, remove the latest from the Order.
                if (pList.Count == 0)
                {
                    Products.Remove(Products.Last());
                }
                else if (pList != null)
                {
                    while (pList.Count > 0)
                        Products.Remove((Product)pList[0]);
                }
            }
        }
        private void RefreshGrid()
        {
            gridProducts.ItemsSource = null;
            gridProducts.ItemsSource = Orders.Last().Products;
        }


        private void gridProducts_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Column.Header.ToString() == "Quantity")
            {
                int newValue = int.Parse(((TextBox)e.EditingElement).Text);
                Product p = gridProducts.SelectedItem as Product;
                p.Quantity = newValue;
                p.Recalculate();
                RefreshGrid();
                Order.CalculateTotal();
            }
        }

        private void btnNewOrder_Click(object sender, RoutedEventArgs e)
        {
            // Create new order
            Order o = new Order();
            o.Number = Order.Number + 1;

            // Add new order to the collection.
            Orders.Add(o);

            // Add default person.
            Person p = Persons[0];
            Order.Recipient = p.Name;
            Order.Address = p.Address;
            Order.City = p.City;
            Order.Phone = p.Phone;
            Order.Date = DateTime.Now.ToShortDateString();
            Order.Time = "10:00";
            Order.OrderTotal = 0;

            OrderNumber = String.Format("Order {0:D3}", Order.Number);
            ProductsHeader = String.Format("Products list - {0}", OrderNumber);

            RefreshGrid();
            tabControl.SelectedIndex = 0;

            MessageBox.Show(String.Format("The Order {0:D3} has been saved to XML and hided.\nNow, please add some products into Order {1:D3}.", o.Number - 1, o.Number));
        }

        private void cmbRecipient_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox combo = sender as ComboBox;
            if (combo.IsDropDownOpen)
            {
                string selectedPersonName = (combo.SelectedItem as Person).Name;

                if (Order.Recipient != selectedPersonName)
                {
                    Person p = GetPerson(selectedPersonName);
                    if (p != null)
                    {
                        Order.Recipient = p.Name;
                        txtAddress.Text = Order.Address = p.Address;
                        txtCity.Text = Order.City = p.City;
                        txtPhone.Text = Order.Phone = p.Phone;
                    }
                }
            }
        }

        private void btnSaveOrder_Click(object sender, RoutedEventArgs e)
        {
            SerializeToXml();
            MessageBox.Show(String.Format("Order {0:D3} has been saved into XML document!", Order.Number));
        }

        private void SerializeToXml()
        {
            // Create XML with data:
            // To get XML, let's use serialization - serialize our classes Orders and Products into XML document.
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(ObservableCollection<Order>));

            using (FileStream fs = new FileStream("Orders.xml", FileMode.Create))
            {
                xmlSerializer.Serialize(fs, Orders);
            }
        }

        private void btnAddRandomProducts_Click(object sender, RoutedEventArgs e)
        {
            Random rnd = new Random();
            int quantity = rnd.Next(3, 10);

            for (int i = 0; i < quantity; i++)
            {
                int randId = rnd.Next(0, PastryItems.Count - 1);

                Product p = new Product();
                p.Name = PastryItems.ElementAt(randId).Key;
                p.Price = PastryItems.ElementAt(randId).Value;
                p.Quantity = rnd.Next(1, 10);
                p.Recalculate();
                Order.Products.Add(p);
            }
        }
        private void cmbTime_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox combo = sender as ComboBox;

            if (combo.IsDropDownOpen)
            {
                string selectedTime = combo.SelectedValue.ToString();

                Order.Time = selectedTime;
            }
        }
        private void btnGenerateReport_Click(object sender, RoutedEventArgs e)
        {
            // 1. Create XML with data.
            // Serialize our classes Orders and Products into XML document.
            SerializeToXml();

            // Show XML with data.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(@"Orders.xml") { UseShellExecute = true });

            // Create the Dataset and read the XML.
            DataSet ds = new DataSet();

            ds.ReadXml(@"Orders.xml");

            // Load the template document.
            string templatePath = @"InvoiceTemplate.docx";
            if (!File.Exists(templatePath))
                templatePath = @"..\..\InvoiceTemplate.docx";

            // Trying to find the assembly "SautinSoft.Document.dll".
            string assemblyPath = @"..\..\..\..\..\..\..\Bin\Net 4.5\SautinSoft.Document.dll";
            if (!File.Exists(assemblyPath))
                assemblyPath = @"..\..\..\..\Bin\Net 4.5\SautinSoft.Document.dll";
            if (!File.Exists(assemblyPath))
                assemblyPath = @"SautinSoft.Document.dll";


            // 1. Load the template using reflection:
            Assembly asm = Assembly.LoadFrom(assemblyPath);
            Type documentCoreClass = asm.GetType("SautinSoft.Document.DocumentCore");
            MethodInfo methodLoad = documentCoreClass.GetMethod("Load", BindingFlags.Public | BindingFlags.Static, null, CallingConventions.Any, new Type[] { typeof(System.String) }, null);
            var dc = methodLoad.Invoke(null, new object[] { templatePath });
            // Without reflection: 
            // DocumentCore dc = DocumentCore.Load(templatePath);

            // 2. Set Maile Merge Options using reflection:
            PropertyInfo propertyMailMerge = documentCoreClass.GetProperty("MailMerge");
            var objMailMerge = propertyMailMerge.GetValue(dc);
            PropertyInfo propertyClearOptions = objMailMerge.GetType().GetProperty("ClearOptions");
            // Because of value of the Enum item "MailMergeClearOptions.RemoveEmptyRanges" is 2.
            propertyClearOptions.SetValue(objMailMerge, 2);
            // Without reflection:
            // dc.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyRanges;

            // 3. Execute Mail Merge using reflection:
            MethodInfo methodExecute = objMailMerge.GetType().GetMethod("Execute", new Type[] { typeof(System.Object) });
            methodExecute.Invoke(objMailMerge, new object[] { ds.Tables["Order"] });
            // Without reflection:
            // dc.MailMerge.Execute(ds.Tables["Order"]);

            // 4. Save the report
            string reportPath = String.Empty;
            if (rbDocx.IsChecked == true)
                reportPath = "Invoices.docx";
            else if (rbRtf.IsChecked == true)
                reportPath = "Invoices.rtf";
            else if (rbHtml.IsChecked == true)
                reportPath = "Invoices.html";
            else
                reportPath = "Invoices.pdf";

            // Using reflection:
            MethodInfo methodSave = documentCoreClass.GetMethod("Save", new Type[] { typeof(System.String) });
            methodSave.Invoke(dc, new object[] { reportPath });
            // Without reflection:            
            //dc.Save(reportPath);

            // 5. Open the report for the demonstration purposes.            
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(reportPath) { UseShellExecute = true });
        }
    }
}

