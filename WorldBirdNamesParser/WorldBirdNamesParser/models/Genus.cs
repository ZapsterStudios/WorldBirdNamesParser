using System;

namespace WorldBirdNamesParser.models
{
    class Genus
    {
        public int ID { get; set; }

        public int FamilyID { get; set; }

        public string Scientific { get; set; }

        public Genus(int ID, int FamilyID, string Scientific)
        {
            this.ID = ID;
            this.FamilyID = FamilyID;
            this.Scientific = Scientific.Replace("'", "\\'");
        }

        public void Output()
        {
            Console.WriteLine($"INSERT INTO specie_genera (`family_id`, `scientific`) VALUES ((@familyOffset + {this.FamilyID}), '{this.Scientific}');");
        }
    }
}
