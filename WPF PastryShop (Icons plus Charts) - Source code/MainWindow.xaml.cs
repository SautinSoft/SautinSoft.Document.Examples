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
using SautinSoft.Document;
using System.Data;
using SautinSoft.Document.MailMerging;
using SautinSoft.Document.Drawing;
using System.Globalization;

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
            // Scan directory for image files.
            Dictionary<string, string> icons = new Dictionary<string, string>();
            foreach (string iconPath in Directory.EnumerateFiles(@"..\..\icons\", @"*.jpg"))
            {
                icons[Path.GetFileNameWithoutExtension(iconPath.ToLower())] = iconPath;

                switch (Path.GetFileName(iconPath))
                {
                    case "Cherry-apple pie.jpg": icons["Cherry/apple/pie".ToLower()] = iconPath; break;
                    case "Dark-milk-white chocolate.jpg": icons["Dark/milk/white chocolate".ToLower()] = iconPath; break;
                    case "Rice-lemon-vanilla pudding.jpg": icons["Rice/lemon/vanilla pudding".ToLower()] = iconPath; break;
                    case "Spice-cake honey-cake.jpg": icons["Spice-cake, honey-cake".ToLower()] = iconPath; break;
                    case "Chocolate-strawberry-vanilla ice cream.jpg": icons["Chocolate/strawberry/vanilla ice cream".ToLower()] = iconPath; break;
                    default: break;
                }
            }



            // 1. Create XML with data.
            // Serialize our classes Orders and Products into XML document.
            SerializeToXml();

            // Show XML with data.
            //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(@"Orders.xml") { UseShellExecute = true });

            // Create the Dataset and read the XML.
            DataSet ds = new DataSet();

            ds.ReadXml(@"Orders.xml");

            // Load the template document.
            string templatePath = @"InvoiceTemplate.docx";
            if (!File.Exists(templatePath))
                templatePath = @"..\..\InvoiceTemplate.docx";

            // 1. Load the template using reflection:
            DocumentCore dc = DocumentCore.Load(templatePath);

            // 2. Set Mail Merge Options using reflection:
            dc.MailMerge.ClearOptions = MailMergeClearOptions.RemoveEmptyRanges;

            // 2.5 Each product will be decorated by appropriate icon.
            dc.MailMerge.FieldMerging += (senderFM, eFM) =>
            {
                // Insert an icon before the product name
                if (eFM.RangeName == "Product" && eFM.FieldName == "Name")
                {
                    eFM.Inlines.Clear();

                    string iconPath;
                    if (icons.TryGetValue(((string)eFM.Value).ToLower(), out iconPath))
                    {
                        eFM.Inlines.Add(new Picture(dc, iconPath) { Layout = new InlineLayout(new SautinSoft.Document.Drawing.Size(30, 30)) });
                        eFM.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.Tab));
                    }

                    eFM.Inlines.Add(new Run(dc, (string)eFM.Value));
                    eFM.Cancel = false;
                }
                // Add the currency sign into "Total" field. 
                // You may change the culture "en-GB" to any desired.
                if (eFM.RangeName == "Order" && eFM.FieldName == "OrderTotal")
                {
                    decimal total = 0;
                    if (Decimal.TryParse((string)eFM.Value, out total))
                    {
                        eFM.Inlines.Clear();
                        eFM.Inlines.Add(new Run(dc, String.Format(new CultureInfo("en-GB"), "{0:C}", total)));
                    }
                }
            };


            // 3. Execute Mail Merge using reflection:
            dc.MailMerge.Execute(ds.Tables["Order"]);

            // 3.5 Add Pie Charts
            if (chkboxPieCharRecipients.IsChecked.Value)
            {
                CreatePieChartsPage(dc);
            }


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

            dc.Save(reportPath);

            // 5. Open the report for the demonstration purposes.            
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(reportPath) { UseShellExecute = true });
        }
        public void CreatePieChartsPage(DocumentCore dc)
        {
            // 1. Add new Section (for the charts) at the beginning of the report.
            Section sectCharts = new Section(dc);
            sectCharts.PageSetup = dc.Sections[0].PageSetup.Clone();
            sectCharts.PageSetup.PageMargins = new PageMargins()
            {
                Left = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point),
                Right = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point),
                Top = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point)
            };

            dc.Sections.Insert(0, sectCharts);

            Paragraph pageHeading = new Paragraph(dc, new Run(dc, "Statistics", new CharacterFormat()
            {
                FontName = "Calibri",
                Size = 42,
                Bold = true,
                UnderlineStyle = UnderlineType.Single,
                Italic = true,
                FontColor = new Color(244, 176, 131),
                AllCaps = true
            }));
            pageHeading.ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Center;
            sectCharts.Blocks.Add(pageHeading);

            CharacterFormat cf = new CharacterFormat()
            {
                FontName = "Calibri",
                Size = 20,
                UnderlineStyle = UnderlineType.Single,
                FontColor = new Color(244, 176, 131),
            };

            Paragraph recipentHeading = new Paragraph(dc, new Run(dc, "Income from the customers in percentages:",
                cf.Clone()));
            sectCharts.Blocks.Add(recipentHeading);

            Paragraph itemHeading = new Paragraph(dc, new Run(dc, "Sales overview - see the best selling product:", cf.Clone()));

            Shape shpItem = new Shape(dc, new FloatingLayout(new HorizontalPosition(5f, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin),
                new VerticalPosition(130f, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new SautinSoft.Document.Drawing.Size(200, 20, LengthUnit.Millimeter)));
            shpItem.Outline.Fill.SetEmpty();
            shpItem.Text.Blocks.Add(itemHeading);
            sectCharts.Content.Start.Insert(shpItem.Content);

            // 2. Add Chart - Recipients
            // Create dictionary for the chart <Recipient Name, Total Percentage>.
            Dictionary<string, double> recipients = new Dictionary<string, double>();
            double totalAmount = Orders.Sum(o => o.OrderTotal);

            foreach (Order o in Orders)
            {
                if (recipients.Keys.Contains(o.Recipient))
                    recipients[o.Recipient] += o.OrderTotal * 100 / totalAmount;
                else
                    recipients.Add(o.Recipient, o.OrderTotal * 100 / totalAmount);
                recipients[o.Recipient] = Math.Round(recipients[o.Recipient]);
            }
            FloatingLayout flR = new FloatingLayout(new HorizontalPosition(20f, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin),
                new VerticalPosition(40f, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new SautinSoft.Document.Drawing.Size(200, 200));

            AddPieChart(sectCharts, flR, recipients, true, "%", true, true);

            // 3. Add Chart - Best selling item
            // Create dictionary for the chart <Product Name, Total>.
            Dictionary<string, double> productSales = new Dictionary<string, double>();
            foreach (Order o in Orders)
                foreach (Product p in o.Products)
                    if (productSales.Keys.Contains(p.Name))
                        productSales[p.Name] += p.TotalPrice;
                    else
                        productSales.Add(p.Name, p.TotalPrice);

            FloatingLayout flP = new FloatingLayout(new HorizontalPosition(20f, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin),
                new VerticalPosition(150f, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new SautinSoft.Document.Drawing.Size(200, 200));

            AddPieChart(sectCharts, flP, productSales, true, "£", true, false);
        }

        /// <summary>
        /// This method creates a pie chart.
        /// </summary>
        /// <param name="dc">Document</param>
        /// <param name="chartLayout">Chart layout</param>
        /// <param name="data">Chart data</param>
        /// <param name="addLabels">Add labels or not</param>
        /// <param name="labelSign">Label sign</param>
        /// <param name="addLegend">Add legend</param>
        /// <remarks>
        /// This method is made specially with open source code. You may change anything you want:<br />
        /// Chart colors, size, position, font name and font size, position of labels (inside or outside the chart), legend position.<br />
        /// If you need any assistance, feel free to email us: support@sautinsoft.com.<br />
        /// </remarks>
        public static void AddPieChart(Section section, FloatingLayout chartLayout, Dictionary<string, double> data, bool addLabels = true, string labelSign = null, bool addLegend = true, bool isLegendHorizontal = true)
        {
            // Assume that our chart can have 10 pies maximally.
            // And we'll use these colors order by order.
            // You may change the colors and their quantity.
            List<string> colors = new List<string>()
            {
                "#70AD47", // light green
                "#4472C4", // blue
                "#FFC000", // yellow
                "#A5A5A5", // grey
                "#ED7D31", // orange
                "#5B9BD5", // light blue
                "#44546A", // blue and grey
                "#C00000", // red
                "#00B050", // green
                "#9933FF"  // purple
            };
            // 1. To create a circle chart, assume that the sum of all values are 100%
            // and calculate the percent for the each value.
            // Translate all data to perce
            double amount = data.Values.Sum();
            List<double> percentages = new List<double>();

            foreach (double v in data.Values)
                percentages.Add(v * 100 / amount);

            // 2. Translate the percentage value of the each pie into degrees.
            // The whole circle is 360 degrees.
            int pies = data.Values.Count;
            List<double> pieDegrees = new List<double>();
            foreach (double p in percentages)
                pieDegrees.Add(p * 360 / 100);

            // 3. Translate degrees to the "Pie" measurement.
            List<double> pieMeasure = new List<double>();
            // Add the start position.
            pieMeasure.Add(0);
            double currentAngle = 0;
            foreach (double pd in pieDegrees)
            {
                currentAngle += pd;
                pieMeasure.Add(480 * currentAngle / 360);
            }

            // 4. Create the pies.
            Shape originalShape = new Shape(section.Document, chartLayout);

            for (int i = 0; i < pies; i++)
            {
                Shape shpPie = originalShape.Clone(true);
                shpPie.Outline.Fill.SetSolid(Color.White);
                shpPie.Outline.Width = 0.5;

                // Generate a random color if we've exceed the number of standard colors
                if (i>=colors.Count)
                {
                    Random rnd = new Random();
                    string randomColor = String.Format("#{0:X}{1:X}{2:X}", rnd.Next(1, 254), rnd.Next(1, 254), rnd.Next(1, 254));
                    colors.Add(randomColor);
                }
                shpPie.Fill.SetSolid(new Color(colors[i]));

                shpPie.Geometry.SetPreset(Figure.Pie);
                shpPie.Geometry.AdjustValues["adj1"] = 45000 * pieMeasure[i];
                shpPie.Geometry.AdjustValues["adj2"] = 45000 * pieMeasure[i + 1];
                (shpPie.Layout as FloatingLayout).WrappingStyle = WrappingStyle.BehindText;

                section.Content.End.Insert(shpPie.Content);
            }

            // 5. Add labels
            if (addLabels)
            {
                // 0.5 ... 1.2 (inside/outside the circle).
                double multiplier = 0.8;
                double radius = chartLayout.Size.Width / 2 * multiplier;
                currentAngle = 0;
                double labelW = 40;
                double labelH = 20;

                for (int i = 0; i < pieDegrees.Count; i++)
                {
                    currentAngle += pieDegrees[i];
                    double middleAngleDegrees = 360 - (currentAngle - pieDegrees[i] / 2);
                    double middleAngleRad = middleAngleDegrees * (Math.PI / 180);

                    // Calculate the (x, y) on the circle.
                    double x = radius * Math.Cos(middleAngleRad);
                    double y = radius * Math.Sin(middleAngleRad);

                    // Correct the position depending of the label size;
                    x -= labelW / 2;
                    y += labelH / 2;

                    HorizontalPosition centerH = new HorizontalPosition(chartLayout.HorizontalPosition.Value + chartLayout.Size.Width / 2, LengthUnit.Point, HorizontalPositionAnchor.LeftMargin);
                    VerticalPosition centerV = new VerticalPosition(chartLayout.VerticalPosition.Value + chartLayout.Size.Height / 2, LengthUnit.Point, VerticalPositionAnchor.TopMargin);

                    HorizontalPosition labelX = new HorizontalPosition(centerH.Value + x, LengthUnit.Point, HorizontalPositionAnchor.LeftMargin);
                    VerticalPosition labelY = new VerticalPosition(centerV.Value - y, LengthUnit.Point, VerticalPositionAnchor.TopMargin);

                    FloatingLayout labelLayout = new FloatingLayout(labelX, labelY, new SautinSoft.Document.Drawing.Size(labelW, labelH));
                    Shape shpLabel = new Shape(section.Document, labelLayout);
                    shpLabel.Outline.Fill.SetEmpty();
                    shpLabel.Text.Blocks.Content.Start.Insert($"{data.Values.ElementAt(i)}{labelSign}",
                        new CharacterFormat()
                        {
                            FontName = "Arial",
                            Size = 10,
                            //FontColor = new Color("#333333")
                            FontColor = new Color("#FFFFFF")
                        });
                    (shpLabel.Text.Blocks[0] as Paragraph).ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Center;
                    section.Content.End.Insert(shpLabel.Content);
                }

                // 6. Add Legend
                if (addLegend)
                {
                    double legendTopMargin = LengthUnitConverter.Convert(isLegendHorizontal?5:0, LengthUnit.Millimeter, LengthUnit.Point);
                    double legendLeftMargin = LengthUnitConverter.Convert(isLegendHorizontal?-10:10, LengthUnit.Millimeter, LengthUnit.Point);
                    HorizontalPosition legendX = new HorizontalPosition(chartLayout.HorizontalPosition.Value + legendLeftMargin + (isLegendHorizontal ? 0 : chartLayout.Size.Width), LengthUnit.Point, HorizontalPositionAnchor.LeftMargin);
                    VerticalPosition legendY = new VerticalPosition(chartLayout.VerticalPosition.Value + legendTopMargin + (isLegendHorizontal ? chartLayout.Size.Height : 0), LengthUnit.Point, VerticalPositionAnchor.TopMargin);
                    double legendW = isLegendHorizontal ? chartLayout.Size.Width * 2 : chartLayout.Size.Width;
                    double legendH = isLegendHorizontal ? 20 : chartLayout.Size.Height+20;

                    FloatingLayout legendLayout = new FloatingLayout(legendX, legendY, new SautinSoft.Document.Drawing.Size(legendW, legendH));
                    Shape shpLegend = new Shape(section.Document, legendLayout);
                    shpLegend.Outline.Fill.SetEmpty();

                    Paragraph pLegend = new Paragraph(section.Document);
                    pLegend.ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Left;

                    for (int i = 0; i < data.Count; i++)
                    {
                        string legendItem = data.Keys.ElementAt(i);

                        // 183 - circle, "Symbol"
                        // 167 - square, "Wingdings"

                        Run marker = new Run(section.Document, (char)167, new CharacterFormat()
                        {
                            FontColor = new Color(colors[i]),
                            FontName = "Wingdings"
                        });
                        pLegend.Content.End.Insert(marker.Content);
                        pLegend.Content.End.Insert($" {legendItem}   ", new CharacterFormat());
                        if (!isLegendHorizontal)
                            pLegend.Inlines.Add(new SpecialCharacter(section.Document, SpecialCharacterType.LineBreak));
                    }


                    shpLegend.Text.Blocks.Add(pLegend);
                    section.Content.End.Insert(shpLegend.Content);
                }
            }
        }

    }
}

