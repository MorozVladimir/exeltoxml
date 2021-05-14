using System;
using System.Collections.Generic;
using System.Text;

namespace exeltoxml.XMLConstruction
{
    public class good
    {
        public int id { get; set; }
        public bool available { get; set; }
        public decimal price { get; set; }
        public decimal priceOld { get; set; }
        public decimal pricePromo { get; set; }
        public decimal stockQuantity { get; set; }
        public string CurrencyId { get; set; }
        public int categoryId { get; set; }
        public List<string> pictures { get; set; }
        public string name { get; set; }
        public string article { get; set; }
        public string vendor { get; set; }
        public string description { get; set; }
        public List<goodParam> parametrs { get; set; }

        public good()
        {
            pictures = new List<string>();
            parametrs = new List<goodParam>();
        }

    }
}
