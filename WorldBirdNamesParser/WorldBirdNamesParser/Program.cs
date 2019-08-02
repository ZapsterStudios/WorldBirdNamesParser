using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace WorldBirdNamesParser
{
    class Program
    {
        static readonly string file = "demo.xlsx";
        static readonly string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        static readonly List<List<string>> columns = new List<List<string>> {
            new List<string> { "scientific", "catalan_name", "czech_name", "estonian_name", "german_name", "indonesian_name", "latvian_name", "norwegian_name", "russian_name", "spanish_name", "ukrainian_name" },
            new List<string> { "english_name", "chinese_name", "danish_name", "finnish_name", "hungarian_name", "italian_name", "lithuanian_name", "polish_name", "slovak_name", "swedish_name" },
            new List<string> { "afrikaans_name", "chinese_traditional_name", "dutch_name", "french_name", "icelandic_name", "japanese_name", "northern_sami_name", "portuguese_name", "slovenian_name", "thai_name" },
        };

        static readonly List<string> excluded = new List<string> {
            "catalan_name", "czech_name", "estonian_name", "indonesian_name", "latvian_name", "ukrainian_name",
            "finnish_name", "hungarian_name", "lithuanian_name", "polish_name", "slovak_name",
            "afrikaans_name", "chinese_traditional_name", "icelandic_name", "northern_sami_name", "slovenian_name", "thai_name",
        };

        static void Main(string[] args)
        {
            // Create new excel application and open file.
            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Open(Path.Combine(path, file));

            // Prepare relation and row ids.
            int rowID = 0;
            int orderID = 1;
            int familyID = 1;

            // Prepare insertion text line.
            string insertion = "INSERT INTO species (`family_id`";
            foreach (List<string> lst in columns)
            {
                foreach (string col in lst)
                {
                    // Add to insertion if not excluded.
                    if (excluded.Contains(col)) continue;
                    insertion += $", `{col}`";
                }
            }

            // Get sheet range and loop through rows.
            Range range = workbook.Sheets[1].UsedRange;
            for (int row = 4; row <= range.Rows.Count; row++)
            {
                // Check if row is order or family.
                if (range.Cells[row, 2].Value2 != null) {
                    // Output order insertion line.
                    Console.WriteLine($"INSERT INTO specie_orders (`class_id`, `scientific`) VALUES (1, '{range.Cells[row, 2].Value2}');");

                    // Up relation id and skip.
                    rowID = 0;
                    orderID++;
                    continue;
                } else if (range.Cells[row, 3].Value2 != null) {
                    // Output family insertion line.
                    Console.WriteLine($"INSERT INTO specie_families (`order_id`, `scientific`) VALUES ({orderID}, '{range.Cells[row, 3].Value2}');");

                    // Up relation id and skip.
                    rowID = 0;
                    familyID++;
                    continue;
                }

                // Output insertion line with family id if first row.
                if (rowID == 0) {
                    Console.Write(insertion + $") VALUES ({familyID}");
                }

                // Loop through and columns.
                for (int col = 4 + rowID; col <= range.Columns.Count - rowID; col += 3)
                {
                    // Skip if column is excluded from output.
                    if (excluded.Contains(columns[rowID][(col - 4 - rowID) / 3])) continue;

                    // Output either null string or value.
                    if (range.Cells[row, col].Value2 == null) {
                        Console.Write(", NULL");
                    } else {
                        Console.Write($", '{range.Cells[row, col].Value2}'");
                    }
                }

                // Increment current row id.
                rowID++;

                // Output ending insertion line if last row.
                if (rowID >= 3) {
                    rowID = 0;
                    Console.WriteLine(");");
                }
            }

            // Wait for user input.
            Console.ReadKey();
        }
    }
}
