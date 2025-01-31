﻿using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
var students = GetStudentList().OrderBy(s => s.Name).ToList(); // Fetch and sort students by name

            var templatePath = "path/to/your/template.xlsx"; // Path to your Excel file A
            var outputPath = "path/to/output/fileB.xlsx"; // Path to save Excel file B

            using (var templatePackage = new ExcelPackage(new FileInfo(templatePath)))
            {
                var templateWorksheet = templatePackage.Workbook.Worksheets[0];

                // Create a new Excel package for output
                using (var outputPackage = new ExcelPackage())
                {
                    var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Sheet1");

                    // Copy structure and styles from template to output
                    templateWorksheet.Cells.Copy(outputWorksheet.Cells);

                    // Replace shortcodes with actual data and additional shortcodes
                    FindAndReplaceShortcode(outputWorksheet, "[title]", "Student List"); // Replace title shortcode
                    FindAndReplaceShortcode(outputWorksheet, "[logo]", "path/to/logo.png"); // Replace logo shortcode

                    foreach (var student in students)
                    {
                        FindAndReplaceShortcode(outputWorksheet, "[studentListData.Name]", student.Name);
                        FindAndReplaceShortcode(outputWorksheet, "[studentListData.Age]", student.Age);
                        FindAndReplaceShortcode(outputWorksheet, "[studentListData.Grade]", student.Grade);
                    }

                    // Fill in the StudentSum shortcode
                    FindAndReplaceShortcode(outputWorksheet, "[StudentSum]", students.Count);

                    // Save the new file
                    outputPackage.SaveAs(new FileInfo(outputPath));
                }
            }

            Console.WriteLine("Excel file B created successfully.");
        }

        static void FindAndReplaceShortcode(ExcelWorksheet worksheet, string shortcode, object value)
        {
            // Search for the shortcode in the worksheet and replace it with the value
            foreach (var cell in worksheet.Cells)
            {
                if (cell.Text.Contains(shortcode))
                {
                    cell.Value = cell.Text.Replace(shortcode, value.ToString());
                }
            }
        }

        static List<Student> GetStudentList()
        {
            // Replace this with your database fetching logic
            return new List<Student>
            {
                new Student { Name = "John Doe", Age = 20, Grade = "A" },
                new Student { Name = "Jane Smith", Age = 22, Grade = "B" }
                // Add more students as needed
            };
        }
    }
}
