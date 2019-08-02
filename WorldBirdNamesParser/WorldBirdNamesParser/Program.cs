using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace WorldBirdNamesParser
{
    class Program
    {
        static readonly string file = "demo.xlsx";
        static readonly string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        static void Main(string[] args)
        {
            // Create new excel application and open file.
            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Open(Path.Combine(path, file));

            // Get sheet range and loop through rows and columns.
            Range range = workbook.Sheets[1].UsedRange;
            for (int row = 1; row <= range.Rows.Count; row++)
            {
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    // Skip if cell is empty.
                    if (range.Cells[row, col].Value2 == null) continue;

                    // Output cell value to console.
                    Console.WriteLine(range.Cells[row, col].Value2);
                }
            }

            // Wait for user input.
            Console.ReadKey();
        }
    }
}
