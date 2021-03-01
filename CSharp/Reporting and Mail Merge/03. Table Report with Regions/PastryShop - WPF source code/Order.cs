using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Windows;
using System.ComponentModel;
using System.Collections.Specialized;

namespace PastryShop
{
    [Serializable]
    public class Order : INotifyPropertyChanged
    {
        public Order()
        {
            Products = new ObservableCollection<Product>();
            OrderTotal = 0;

            Products.CollectionChanged += (sender, e) =>
            {
                if (e.Action == NotifyCollectionChangedAction.Add || e.Action==NotifyCollectionChangedAction.Remove)
                {
                    CalculateTotal();
                }
            };
        }
        
        private string _date;
        private int _number;
        private string _recipient;
        private string _address;
        private string _city;
        private string _phone;
        private string _time;
        private double _orderTotal;




        public string Date
        {
            get { return _date; }
            set
            {
                if (value != _date)
                {
                    _date = value;
                    OnPropertyChanged("Date");
                }
            }
        }
        public int Number
        {
            get { return _number; }
            set
            {
                if (value != _number)
                {
                    _number = value;
                    OnPropertyChanged("Number");
                }
            }
        }
        public string Recipient
        {
            get { return _recipient; }
            set
            {
                if (value != _recipient)
                {
                    _recipient = value;
                    OnPropertyChanged("Recipient");
                }
            }
        }
        public string Address
        {
            get { return _address; }

            set
            {
                if (value != _address)
                {
                    _address = value;
                    OnPropertyChanged("Address");
                }
            }
        }
        public string City
        {
            get { return _city; }
            set
            {
                if (value != _city)
                {
                    _city = value;
                    OnPropertyChanged("City");
                }
            }
        }
        public string Phone
        {
            get { return _phone; }
            set
            {
                if (value != _phone)
                {
                    _phone = value;
                    OnPropertyChanged("Phone");
                }
            }
        }
        public string Time
        {
            get { return _time; }
            set
            {
                if (value != _time)
                {
                    _time = value;
                    OnPropertyChanged("Time");
                }
            }
        }
        public double OrderTotal
        {
            get { return _orderTotal; }
            set
            {
                if (value != _orderTotal)
                {
                    _orderTotal = value;                    
                }
                OnPropertyChanged("OrderTotal");
            }
        }

        [XmlElement("Product")]
        public ObservableCollection<Product> Products { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void CalculateTotal()
        {
            // Calculate total
            double total = 0;
            foreach (Product p in Products)
                total += p.TotalPrice;

            OrderTotal = total;
        }
    }
}
