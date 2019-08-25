using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorldBirdNamesParser.models
{
    class Order
    {
        public int ID { get; set; }

        public int ClassID { get; set; }

        public string Scientific { get; set; }

        public Order(int ID, int ClassID, string Scientific)
        {
            this.ID = ID;
            this.ClassID = ClassID;
            this.Scientific = Scientific.Replace("'", "\\'");
        }

        public void Output()
        {
            Console.WriteLine($"INSERT INTO specie_orders (`id`, `class_id`, `scientific`) VALUES ({this.ID}, {this.ClassID}, '{this.Scientific}');");
        }
    }
}
