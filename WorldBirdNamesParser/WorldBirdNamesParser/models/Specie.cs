using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorldBirdNamesParser.models
{
    class Specie
    {
        public int ID { get; set; }

        public int FamilyID { get; set; }

        public string SpecieID { get; set; }

        public string Scientific { get; set; }

        public List<KeyValuePair<string, string>> Locales { get; } = new List<KeyValuePair<string, string>>();

        public Specie(int ID, int FamilyID, int SpecieID, string Scientific)
        {
            this.ID = ID;
            this.FamilyID = FamilyID;
            this.SpecieID = (SpecieID == 0 ? "NULL" : ""+SpecieID);
            this.Scientific = Scientific.Replace("'", "\\'");
        }

        public void Output()
        {
            Console.Write("INSERT INTO species (`id`, `family_id`, `specie_id`, `scientific`");

            foreach (KeyValuePair<string, string> pair in this.Locales)
            {
                Console.Write($", `{pair.Key}`");
            }

            Console.Write($") VALUES ({this.ID}, {this.FamilyID}, {this.SpecieID}, '{this.Scientific}'");
            
            foreach (KeyValuePair<string, string> pair in this.Locales)
            {
                Console.Write($", '{pair.Value.Replace("'", "\\'")}'");
            }
            
            Console.WriteLine(");");
        }
    }
}
