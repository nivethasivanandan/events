using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;
using System;
using System.ComponentModel;
using System.IO;

class Program
{
    static void Main()
    {
        try
        {
            

            // Hardcoded dataset
            var dataset = new object[,]
            {
                {"Name", "Age", "Salary", "Education", "Risk"},
                {"Bob", 30, 10000, "BE", "High"},
                {"Alice", 32, 30000, "BE", "High"},
                {"Antony", 32, 30000, "BE", "Low"},
                {"Charles", 34, 10000, "BE", "Critical"},
                {"Kevin", 30, 40000, "BE", "Medium"},
                {"Roshan", 30, 10000, "BE", "Medium"},
                {"Sam", 31, 20000, "BE", "Low"},
                {"Jack", 45, 20000, "ME", "Low"},
                {"Jim", 34, 20000, "ME", "Critical"},
                {"Mark", 46, 30000, "ME", "Critical"}
            };

            // Load the hardcoded dataset into the worksheet
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                for (int row = 1; row <= dataset.GetLength(0); row++)
                {
                    for (int col = 1; col <= dataset.GetLength(1); col++)
                    {
                        worksheet.Cells[row, col].Value = dataset[row - 1, col - 1];
                    }
                }

                // Display data read from the worksheet
                Console.WriteLine("Data Read from Worksheet:");
                DisplayData(worksheet);

                // Create a pie chart for the original data
                CreatePieChart(worksheet, "OriginalPieChart");

                // Create a bar chart for the original data
                CreateBarChart(worksheet, "OriginalBarChart");

                // Create a column clustered chart for the original data
                CreateColumnClusteredChart(worksheet, "OriginalColumnClusteredChart");

                // Read data from the worksheet
                // Modify the data as needed
                // Example: Add a new value to cell A1
                worksheet.Cells["A1"].Value = "New Value";

                // Clean and convert Salary values
                CleanAndConvertSalary(worksheet);

                // Save the changes to the output Excel file
                string outputFile = "C:\\Users\\Administrator\\Desktop\\epplus_output.xlsx";
                package.SaveAs(new FileInfo(outputFile));

                // Display modified data
                Console.WriteLine("\nModified Data:");
                DisplayData(worksheet);

                // Create a pivot table
                var pivotTable = worksheet.Tables.Add(worksheet.Cells["A1:E11"], "MyPivotTable");
                pivotTable.TableStyle = TableStyles.Medium2;

                // Save the changes to the output Excel file
                package.SaveAs(new FileInfo(outputFile));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static void DisplayData(ExcelWorksheet worksheet)
    {
        for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                Console.Write($"{worksheet.Cells[row, col].Text}\t");
            }
            Console.WriteLine();
        }
    }

    static void CleanAndConvertSalary(ExcelWorksheet worksheet)
    {
        // Assuming Salary values are in column C (3rd column)
        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
        {
            // Clean and convert Salary value to numeric
            string salaryText = worksheet.Cells[row, 3].Text.Replace(",", "");
            if (double.TryParse(salaryText, out double salaryNumeric))
            {
                worksheet.Cells[row, 3].Value = salaryNumeric;
            }
            else
            {
                Console.WriteLine($"Unable to convert Salary value at row {row} to numeric.");
            }
        }
    }

    static void CreatePieChart(ExcelWorksheet worksheet, string chartName)
    {
        // Display chart name for debugging
        Console.WriteLine($"Creating {chartName}...");

        // Create a pie chart
        var chart = worksheet.Drawings.AddChart(chartName, eChartType.Pie);
        chart.SetPosition(1, 0, 10, 0);
        chart.SetSize(600, 400);
        chart.Title.Text = $"{chartName} Chart";

        // Set data range for the chart
        var dataRange = worksheet.Cells["C2:C11"];
        var labelsRange = worksheet.Cells["A2:A11"];
        var series = chart.Series.Add(dataRange, labelsRange);

        // Add a legend for better visibility
        chart.Legend.Add();
    }

    static void CreateBarChart(ExcelWorksheet worksheet, string chartName)
    {
        // Display chart name for debugging
        Console.WriteLine($"Creating {chartName}...");

        // Create a bar chart
        var chart = worksheet.Drawings.AddChart(chartName, eChartType.BarClustered);
        chart.SetPosition(30, 0, 10, 0);
        chart.SetSize(600, 400);
        chart.Title.Text = $"{chartName} Chart";

        // Set data range for the chart
        var dataRange = worksheet.Cells["C2:C11"];
        var labelsRange = worksheet.Cells["A2:A11"];
        var series = chart.Series.Add(dataRange, labelsRange);

        // Add a legend for better visibility
        chart.Legend.Add();
    }
    static void CreateColumnClusteredChart(ExcelWorksheet worksheet, string chartName)
    {
        // Display chart name for debugging
        Console.WriteLine($"Creating {chartName}...");

        // Create a column clustered chart
        var chart = worksheet.Drawings.AddChart(chartName, eChartType.ColumnClustered);
        chart.SetPosition(60, 0, 10, 0);
        chart.SetSize(600, 400);
        chart.Title.Text = $"{chartName} Chart";

        // Set data range for the chart
        var dataRange = worksheet.Cells["C2:C11"];
        var labelsRange = worksheet.Cells["A2:A11"];
        var series = chart.Series.Add(dataRange, labelsRange);

        // Add a legend for better visibility
        chart.Legend.Add();
    }


}
