using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Xml;
using WorldBirdNamesParser.models;

namespace WorldBirdNamesParser
{
    class Program
    {
        readonly List<Order> Orders = new List<Order>();
        readonly List<Family> Families = new List<Family>();
        readonly List<Specie> Species = new List<Specie>();

        readonly List<List<string>> locales = new List<List<string>> {
            new List<string> { "scientific", "catalan_name", "czech_name", "estonian_name", "german_name", "indonesian_name", "latvian_name", "norwegian_name", "russian_name", "spanish_name", "ukrainian_name" },
            new List<string> { "english_name", "chinese_name", "danish_name", "finnish_name", "hungarian_name", "italian_name", "lithuanian_name", "polish_name", "slovak_name", "swedish_name" },
            new List<string> { "afrikaans_name", "chinese_traditional_name", "dutch_name", "french_name", "icelandic_name", "japanese_name", "northern_sami_name", "portuguese_name", "slovenian_name", "thai_name" },
        };

        readonly List<string> excluded = new List<string> {
            "scientific", "catalan_name", "czech_name", "estonian_name", "indonesian_name", "latvian_name", "ukrainian_name",
            "finnish_name", "hungarian_name", "lithuanian_name", "polish_name", "slovak_name",
            "afrikaans_name", "chinese_traditional_name", "icelandic_name", "northern_sami_name", "slovenian_name", "thai_name",
        };

        static void Main(string[] args)
        {
            // Make console use UTF-8 for outputting.
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            // Start program processing.
            Program program = new Program();
            program.ProcessBase(args[0]);
            program.ProcessLocale(args[1]);
            program.Output();
        }

        void ProcessBase(string file)
        {
            // Output starting messages.
            Console.WriteLine("-- Starting base processing");

            // Create reader and load XML file.
            XmlDocument document = new XmlDocument();
            document.Load(file);

            // Loop through all order nodes.
            int count = document.GetElementsByTagName("order").Count;
            foreach (XmlNode order in document.GetElementsByTagName("order"))
            {
                // Generate id for order.
                int orderID = this.Orders.Count + 1;

                // Get name and uppercase first name only.
                string orderName = order.SelectSingleNode("latin_name").InnerText.ToLower();
                orderName = orderName.Substring(0, 1).ToUpper() + orderName.Substring(1);

                // Add new order to list.
                this.Orders.Add(new Order(orderID, 1, orderName));

                // Output status message.
                Console.WriteLine($"-- Processed {orderID} of {count} orders");

                // Loop through all family nodes.
                foreach (XmlNode family in order.SelectNodes("family"))
                {
                    // Generate id for family.
                    int familyID = this.Families.Count + 1;

                    // Get family latin and english name.
                    string familyName = family.SelectSingleNode("latin_name").InnerText;
                    string englishFamilyName = family.SelectSingleNode("english_name").InnerText;

                    // Add new family to list.
                    this.Families.Add(new Family(familyID, orderID, familyName, englishFamilyName));

                    // Loop through all genus nodes.
                    foreach (XmlNode genus in family.SelectNodes("genus"))
                    {
                        // Get genus prefix name.
                        string genusName = genus.SelectSingleNode("latin_name").InnerText;

                        // Loop through all specie nodes.
                        foreach (XmlNode specie in genus.SelectNodes("species"))
                        {
                            // Generate id for specie.
                            int specieID = this.Species.Count + 1;

                            // Get specie latin name and combine with genus name.
                            string shortSpecieName = specie.SelectSingleNode("latin_name").InnerText;
                            string fullSpecieName = genusName + " " + shortSpecieName;

                            // Get specie english name.
                            string englishSpecieName = specie.SelectSingleNode("english_name").InnerText;

                            // Add new specie to list.
                            this.Species.Add(new Specie(specieID, familyID, 0, fullSpecieName, englishSpecieName));

                            // Loop through all sub-specie nodes.
                            foreach (XmlNode subspecie in specie.SelectNodes("subspecies"))
                            {
                                // Generate subspecie id and get latin name.
                                int subSpecieID = this.Species.Count + 1;
                                string subSpecieName = subspecie.SelectSingleNode("latin_name").InnerText;

                                // Add new sub-specie to list.
                                this.Species.Add(new Specie(subSpecieID, familyID, specieID, fullSpecieName + " " + subSpecieName, null));
                            }
                        }
                    }

                }
            }
        }

        void ProcessLocale(string file)
        {
            // Output starting messages.
            Console.WriteLine("-- Starting locale processing");

            // Create new excel application and open file.
            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Open(file);

            // Define row counter and specie holder.
            int rowID = 0;
            int rowCounter = 0;
            Specie specie = null;

            // Output UTF-8 encoding and aves class.
            Console.WriteLine("SET NAMES utf8;");
            Console.WriteLine("INSERT INTO specie_classes (`scientific`) VALUES ('Aves');");

            // Get sheet range and loop through rows.
            Range range = workbook.Sheets[1].UsedRange;
            for (int row = 4; row <= range.Rows.Count; row++)
            {
                // Skip order and family rows.
                if (range.Cells[row, 2].Value2 != null || range.Cells[row, 3].Value2 != null)
                {
                    continue;
                }

                // Loop through rows and columns.
                for (int col = 4 + rowID; col <= range.Columns.Count - rowID; col += 3)
                {
                    // Get the cell name and value.
                    string name = this.locales[rowID][(col - 4 - rowID) / 3];
                    string value = Convert.ToString(range.Cells[row, col].Value2);
                    
                    // Check if at new scientific row.
                    if (name.Equals("scientific"))
                    {
                        // Find specie for found scientific name.
                        specie = this.Species.Find(item => item.Scientific == value);

                        // Output error and kill program if missing.
                        if (specie == null)
                        {
                            Console.WriteLine("Unable to find specie: " + value);
                            Environment.Exit(1);
                        }
                    }

                    // Skip if name is excluded.
                    if (this.excluded.Contains(name)) continue;

                    // Add new locale to specie.
                    if (value != null && value != "") {
                        specie.Locales.Add(new KeyValuePair<string, string>(name, value));
                    }
                }

                // Increment current row id.
                rowID++;
                rowCounter++;

                // Reset row id once completed.
                if (rowID >= 3)
                {
                    rowID = 0;
                }

                // Output status message.
                if ((rowCounter % 100) == 0)
                {
                    rowCounter = 0;
                    Console.WriteLine($"-- Processed {row} of {range.Rows.Count} rows");
                }
            }

            // Close worksbook and exit excel.
            workbook.Close();
            excel.Quit();
        }

        void Output()
        {
            // Output starting messages.
            Console.WriteLine("-- OUTPUT --");

            // Output charset and default class.
            Console.WriteLine("SET NAMES utf8;");
            Console.WriteLine("INSERT INTO specie_classes (`scientific`) VALUES ('Aves');");

            // Output the order, family, and specie rows.
            foreach (Order order in this.Orders) order.Output();
            foreach (Family family in this.Families) family.Output();
            foreach (Specie specie in this.Species) specie.Output();
        }
    }
}
