using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorldBirdNamesParser.models
{
    class Family
    {
        public int ID { get; set; }

        public int OrderID { get; set; }

        public string Scientific { get; set; }

        public string English { get; set; }

        public Family(int ID, int OrderID, string Scientific, string English = null)
        {
            this.ID = ID;
            this.OrderID = OrderID;
            this.Scientific = Scientific.Replace("'", "\\'");
            this.English = English.Replace("'", "\\'");
        }

        public void Output()
        {
            Console.WriteLine($"INSERT INTO specie_families (`id`, `order_id`, `scientific`, `english_name`) VALUES ({this.ID}, {this.OrderID}, '{this.Scientific}', '{this.English}');");
        }
    }
}
