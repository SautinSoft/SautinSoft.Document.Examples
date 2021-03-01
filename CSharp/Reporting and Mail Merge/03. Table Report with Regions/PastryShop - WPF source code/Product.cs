using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace PastryShop
{
    [Serializable]
    public class Product
    {
        public string Name { get; set; }        
        public double Price { get; set; }
        public int Quantity
        {
            get;
            set;
        }
        public double TotalPrice { get; set; }
        public Product() { }
        public Product(string name, double price, int quantity)
        {
            Name = name;            
            Price = price;
            Quantity = quantity;
            Recalculate();
        }
        public void Recalculate()
        {
            TotalPrice = Quantity * Price;
        }
    }
}
