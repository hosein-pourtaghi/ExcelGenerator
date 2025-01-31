using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var students = GetStudentList(); // Fetch from your database
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

                    // Replace shortcodes with actual data
                    ReplaceShortcodes(outputWorksheet, students);

                    // Save the new file
                    outputPackage.SaveAs(new FileInfo(outputPath));
                }
            }

            Console.WriteLine("Excel file B created successfully.");
        }

        static void ReplaceShortcodes(ExcelWorksheet worksheet, List<Student> students)
        {
            // Example shortcode replacements
            worksheet.Cells["A1"].Value = "Your Logo Here"; // Replace [logo]
            worksheet.Cells["B1"].Value = "Your Title Here"; // Replace [title]

            // Assuming student data starts at row 3
            int startRow = 3;
            foreach (var student in students)
            {
                worksheet.Cells[startRow, 1].Value = student.Name; // Replace [studentListData.Name]
                worksheet.Cells[startRow, 2].Value = student.Age; // Replace [studentListData.Age]
                worksheet.Cells[startRow, 3].Value = student.Grade; // Replace [studentListData.Grade]
                startRow++;
            }

            // Fill in the StudentSum shortcode
            worksheet.Cells[startRow, 1].Value = "Total Students:"; // Label for total
            worksheet.Cells[startRow, 2].Value = students.Count; // Replace [StudentSum]
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
