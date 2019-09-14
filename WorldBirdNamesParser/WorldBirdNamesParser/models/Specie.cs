using System;
using System.Collections.Generic;

namespace WorldBirdNamesParser.models
{
    class Specie
    {
        public int ID { get; set; }

        public int GenusID { get; set; }

        public string SpecieID { get; set; }

        public string Scientific { get; set; }

        public List<KeyValuePair<string, string>> Locales { get; } = new List<KeyValuePair<string, string>>();

        public Specie(int ID, int GenusID, int SpecieID, string Scientific)
        {
            this.ID = ID;
            this.GenusID = GenusID;
            this.SpecieID = (SpecieID == 0 ? "NULL" : ""+SpecieID);
            this.Scientific = Scientific.Replace("'", "\\'");
        }

        public void Output()
        {
            Console.Write("INSERT INTO species (`genus_id`, `specie_id`, `scientific`");

            foreach (KeyValuePair<string, string> pair in this.Locales)
            {
                Console.Write($", `{pair.Key}`");
            }

            Console.Write($") VALUES ((@genusOffset + {this.GenusID}), (@specieOffset + {this.SpecieID}), '{this.Scientific}'");
            
            foreach (KeyValuePair<string, string> pair in this.Locales)
            {
                Console.Write($", '{pair.Value.Replace("'", "\\'")}'");
            }
            
            Console.WriteLine(");");
        }
    }
}
