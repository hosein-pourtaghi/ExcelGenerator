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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set the license context

            var students = GetStudentList().OrderBy(s => s.Name).ToList(); // Fetch and sort students by name

            var templatePath = "./template.xlsx"; // Path to your Excel file A
            var outputPath = "../fileB.xlsx"; // Path to save Excel file B

            using (var templatePackage = new ExcelPackage(new FileInfo(templatePath)))
            {
                var templateWorksheet = templatePackage.Workbook.Worksheets[0];
                Console.WriteLine($"Number of worksheets: {templatePackage.Workbook.Worksheets.Count}");


                // Log the names of all available worksheets
                foreach (var sheet in templatePackage.Workbook.Worksheets)
                {
                    Console.WriteLine($"Worksheet name: {sheet.Name}");
                }

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
                        // Insert student data into the worksheet
                        int row = 2; // Start from the second row (assuming the first row is for headers) 
                        foreach (var s in students)
                        {
                            outputWorksheet.Cells[row, 1].Value = s.Name; // Column 1 for Name
                            outputWorksheet.Cells[row, 2].Value = s.Age;  // Column 2 for Age
                            outputWorksheet.Cells[row, 3].Value = s.Grade; // Column 3 for Grade
                            row++;
                        }
                    }

                    // Fill in the StudentSum shortcode
                    // FindAndReplaceShortcode(outputWorksheet, "[StudentSum]", students.Count);
                    Console.WriteLine("Excel file B finished.");

                    try
                    {
                        // Save the new file
                        outputPackage.SaveAs(new FileInfo(outputPath));
                        Console.WriteLine("Excel file B saved.");

                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("exception occur", e.Message);
                    }

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
                new Student { Name = "John Doe", Age = 2, Grade = "A" },
                new Student { Name = "John Doe", Age = 23, Grade = "A" },
                new Student { Name = "John Doe", Age = 24, Grade = "A" },
                new Student { Name = "John Doe", Age = 25, Grade = "A" },
                new Student { Name = "John Doe", Age = 26, Grade = "A" },
                new Student { Name = "John Doe", Age = 27, Grade = "A" },
                new Student { Name = "John Doe", Age = 21, Grade = "A" },
                new Student { Name = "John Doe", Age = 27, Grade = "A" },
                new Student { Name = "John Doe", Age = 26, Grade = "A" },
                new Student { Name = "Jane Smith", Age = 22, Grade = "B" }
                // Add more students as needed
            };
        }
    }
}
