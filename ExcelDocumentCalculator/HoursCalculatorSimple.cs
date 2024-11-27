using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization; 

namespace ExcelDocumentCalculator
{
    public class HoursCalculatorSimple
    {
        public void Calculate(string inputFilePath, string invoiceTemplatePath, Action callBack, int maxWritesInInvoice, int moneyMinLimit, int moneyMaxLimit, float hourRate)
        {
            string directoryPath = Path.GetDirectoryName(inputFilePath);
            string archiveFilePath = Path.Combine(directoryPath, $"Archive_{DateTime.Now:yyyy_MM_dd}.xlsx");
            string newWorkingHoursFilePath = Path.Combine(directoryPath, $"NewWorkingHours_{DateTime.Now:yyyy_MM_dd}.xlsx");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(inputFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Assuming the data is in the first worksheet
                var rows = worksheet.Dimension.Rows;
                var columns = worksheet.Dimension.Columns;

                // Find column indexes dynamically
                var headers = new Dictionary<string, int>();
                for (int col = 1; col <= columns; col++)
                {
                    string header = worksheet.Cells[1, col].Text.Trim().ToLower();
                    headers[header] = col;
                }

                // Validate required columns
                if (!headers.ContainsKey("date") || !headers.ContainsKey("hours") ||
                    !headers.ContainsKey("projectid") || !headers.ContainsKey("workdescription"))
                {
                    Console.WriteLine("Missing one or more required columns: Date, Hours, ProjectId, WorkDescription.");
                    return;
                }

                int dateCol = headers["date"];
                int hoursCol = headers["hours"];
                int projectIdCol = headers["projectid"];
                int workDescriptionCol = headers["workdescription"];
                int ignoreCol = headers.ContainsKey("ignore") ? headers["ignore"] : -1; // Optional column

                var projectData = new Dictionary<string, ProjectSummary>();
                var rowsToKeep = new List<Tuple<int, DateTime>>();

                for (int i = 2; i <= rows; i++) // Start from 2 to skip the header row
                {
                    // Check for Ignore column
                    if (ignoreCol > 0 && !string.IsNullOrWhiteSpace(worksheet.Cells[i, ignoreCol].Text))
                    {
                        DateTime keepDate = DateTime.Parse(worksheet.Cells[i, dateCol].Text, CultureInfo.InvariantCulture);
                        var index = i;
                        rowsToKeep.Add(Tuple.Create(index, keepDate));
                        continue; // Skip this row for calculation
                    }

                    // Check for partially empty rows
                    if (string.IsNullOrWhiteSpace(worksheet.Cells[i, dateCol].Text) ||
                        string.IsNullOrWhiteSpace(worksheet.Cells[i, hoursCol].Text) ||
                        string.IsNullOrWhiteSpace(worksheet.Cells[i, projectIdCol].Text) ||
                        string.IsNullOrWhiteSpace(worksheet.Cells[i, workDescriptionCol].Text))
                    {
                        continue; // Skip this row
                    }

                    // Parse the data
                    DateTime date = DateTime.Parse(worksheet.Cells[i, dateCol].Text);
                    double hours = double.Parse(worksheet.Cells[i, hoursCol].Text);
                    string projectId = worksheet.Cells[i, projectIdCol].Text;
                    string comment = worksheet.Cells[i, workDescriptionCol].Text;

                    if (!projectData.ContainsKey(projectId))
                    {
                        projectData[projectId] = new ProjectSummary
                        {
                            ProjectId = projectId,
                            DateFrom = date,
                            DateTo = date,
                            TotalHours = 0,
                            Comments = new HashSet<string>(),
                            Dates = new List<DateTime>()
                        };
                    }

                    var project = projectData[projectId];
                    project.TotalHours += hours;
                    project.DateFrom = project.DateFrom > date ? date : project.DateFrom;
                    project.DateTo = project.DateTo < date ? date : project.DateTo;
                    project.Comments.Add(comment);
                    project.Dates.Add(date);
                }


                // Generate invoices
                List<ProjectSummary> unallocatedProjects;
                var invoices = GenerateInvoices(projectData.Values.ToList(), maxWritesInInvoice, moneyMinLimit, moneyMaxLimit, hourRate, out unallocatedProjects);

                CreateInvoiceSummaries(invoiceTemplatePath, invoices, directoryPath);

                // Process invoices (can save to files or use for further logic)
                Console.WriteLine("Generated Invoices:");
                for (int i = 0; i < invoices.Count; i++)
                {
                    Console.WriteLine($"Invoice {i + 1}:");
                    foreach (var project in invoices[i])
                    {
                        Console.WriteLine($"  {project.ProjectId} - {project.TotalHours} hours");
                    }
                }

                foreach (var project in unallocatedProjects)
                {
                    string projectId = project.ProjectId;

                    // Find all rows in the input file that match this projectId
                    for (int i = 2; i <= rows; i++) // Skip header row
                    {
                        if (!rowsToKeep.Any(x=>x.Item1 == i) &&
                            worksheet.Cells[i, projectIdCol].Text == projectId &&
                            (ignoreCol <= 0 || string.IsNullOrWhiteSpace(worksheet.Cells[i, ignoreCol].Text)))
                        {
                            // Parse the data
                            var date = DateTime.Parse(worksheet.Cells[i, dateCol].Text, CultureInfo.InvariantCulture);
                            var index = i;
                            rowsToKeep.Add(Tuple.Create(index, date));
                        }
                    }
                }

                // Create a copy of the input file as an archive
                File.Copy(inputFilePath, archiveFilePath, true);

                
                // Copy the input file to create a new working hours file
                File.Copy(inputFilePath, newWorkingHoursFilePath, true);  // Overwrite if exists

                // Open the copied file using ExcelPackage
                using (var newWorkingHoursPackage = new ExcelPackage(new FileInfo(newWorkingHoursFilePath)))
                {
                    var newWoringHoursSheet = newWorkingHoursPackage.Workbook.Worksheets[0];  // Assuming we're working with the first worksheet
                    var totalRows = newWoringHoursSheet.Dimension.Rows;
                    var totalColumns = newWoringHoursSheet.Dimension.Columns;

                    // Clear all the cells (starting from row 2 to skip headers)
                    for (int row = 2; row <= totalRows; row++)  // Skipping the header row
                    {
                        for (int col = 1; col <= totalColumns; col++)
                        {
                            newWoringHoursSheet.Cells[row, col].Clear();  // Clears the cell content, but retains the styles
                        }
                    }

                    // Now, fill the required cells with the new data
                    int currentRow = 2;
                    rowsToKeep = rowsToKeep.OrderBy(x => x.Item2).ToList();
                    foreach (var keyValue in rowsToKeep)  // 'rowsToKeep' should hold the rows you want to keep
                    {
                        for (int col = 1; col <= totalColumns; col++)
                        {
                            // Copy the value from the original file to the new file (if applicable)
                            newWoringHoursSheet.Cells[currentRow, col].Value = worksheet.Cells[keyValue.Item1, col].Text;
                        }

                        // Optionally clear any specific columns if needed (like "Ignore" column)
                        if (ignoreCol > 0)
                        {
                            newWoringHoursSheet.Cells[currentRow, ignoreCol].Value = string.Empty;
                        }

                        currentRow++;
                    }

                    // Save the updated file
                    newWorkingHoursPackage.Save();
                }

                Console.WriteLine($"New working hours file saved at: {newWorkingHoursFilePath}");

                Console.WriteLine("Summary created, archive saved, and new working hours updated.");
                Console.WriteLine($"Archive: {archiveFilePath}");
                Console.WriteLine($"NewWorkingHours: {newWorkingHoursFilePath}");

                callBack?.Invoke();
            }
        }

        public void WriteInvoiceToTemplate(string templateFilePath, List<ProjectSummary> invoiceProjects, string outputDirectory, string invoiceNumber)
        {
            // Make a copy of the template file with a new name
            string invoiceFilePath = Path.Combine(outputDirectory, $"Invoice_{invoiceNumber}_Summary_{DateTime.Now:yyyy_MM_dd}.xlsx");
            File.Copy(templateFilePath, invoiceFilePath, true);

            // Open the copied file using EPPlus
            using (var package = new ExcelPackage(new FileInfo(invoiceFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0]; // Assuming the data is on the first sheet

                int startRow = 13; // Start from row 13
                int currentRow = startRow;

                foreach (var project in invoiceProjects)
                {
                    // Write project description in Column D (4th column)
                    string projectDescription = $"{project.ProjectId} - {string.Join(", ", project.Comments)} ({GroupDatesIntoPeriods(project.Dates)})";
                    worksheet.Cells[currentRow, 4].Value = projectDescription;

                    // Write project hours in Column I (9th column)
                    worksheet.Cells[currentRow, 9].Value = project.TotalHours;

                    currentRow++;

                    // Stop if we've filled up the available rows (10 rows for working hours)
                    if (currentRow > startRow + 9)
                        break;
                }

                // Save the updated invoice file
                package.Save();
            }

            Console.WriteLine($"Invoice saved: {invoiceFilePath}");
        }

        private List<List<ProjectSummary>> GenerateInvoices(
            List<ProjectSummary> projects,
            int maxWritesInInvoice,
            double moneyMinLimit,
            double moneyMaxLimit,
            double hourRate,
            out List<ProjectSummary> unallocatedProjects)
        {
            var invoices = new List<List<ProjectSummary>>();
            unallocatedProjects = new List<ProjectSummary>();

            var sortedProjects = projects.OrderByDescending(p => p.TotalHours).ToList();

            while (sortedProjects.Count > 0)
            {
                var currentInvoice = new List<ProjectSummary>();
                double currentInvoiceValue = 0;

                for (int i = 0; i < sortedProjects.Count; i++)
                {
                    var project = sortedProjects[i];
                    double projectValue = project.TotalHours * hourRate;

                    if (currentInvoice.Count < maxWritesInInvoice && currentInvoiceValue + projectValue <= moneyMaxLimit)
                    {
                        currentInvoice.Add(project);
                        currentInvoiceValue += projectValue;
                        sortedProjects.RemoveAt(i);
                        i--; // Adjust index after removal
                    }
                }

                if (currentInvoiceValue >= moneyMinLimit)
                {
                    invoices.Add(currentInvoice);
                }
                else
                {
                    // Add projects back to unallocated if not fitting the invoice
                    foreach (var project in currentInvoice)
                        unallocatedProjects.Add(project);

                    // Stop processing further if no valid invoices can be formed
                    break;
                }
            }

            // Add remaining projects that couldn't be allocated
            unallocatedProjects.AddRange(sortedProjects);

            return invoices;
        }

        private void CreateInvoiceSummaries(
            string invoiceTemplatePath,
            List<List<ProjectSummary>> invoices,
            string outputDirectory)
        {
            int invoiceNumber = 1;

            foreach (var invoice in invoices)
            {
                WriteInvoiceToTemplate(invoiceTemplatePath, invoice, outputDirectory, invoiceNumber.ToString());
             
                invoiceNumber++;
            }
        }

        private string GroupDatesIntoPeriods(List<DateTime> dates, int maxGapDays = 7)
        {
            if (dates == null || dates.Count == 0)
                return string.Empty;

            dates = dates.OrderBy(d => d).ToList();
            var groupedDates = new List<string>();
            DateTime? periodStart = null;
            DateTime? periodEnd = null;

            foreach (var date in dates)
            {
                if (periodStart == null)
                {
                    periodStart = date;
                    periodEnd = date;
                }
                else if ((date - periodEnd.Value).TotalDays <= maxGapDays)
                {
                    // Extend the period
                    periodEnd = date;
                }
                else
                {
                    // Add the current period to the list
                    groupedDates.Add(periodStart == periodEnd
                        ? periodStart.Value.ToString("dd.MM.yyyy")
                        : $"{periodStart:dd.MM.yyyy}-{periodEnd:dd.MM.yyyy}");

                    // Start a new period
                    periodStart = date;
                    periodEnd = date;
                }
            }

            // Add the last period
            if (periodStart != null)
            {
                groupedDates.Add(periodStart == periodEnd
                    ? periodStart.Value.ToString("dd.MM.yyyy")
                    : $"{periodStart:dd.MM.yyyy}-{periodEnd:dd.MM.yyyy}");
            }

            return string.Join(", ", groupedDates);
        }
    }
}