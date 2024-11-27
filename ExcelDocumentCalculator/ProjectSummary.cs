using System;
using System.Collections.Generic;

namespace ExcelDocumentCalculator
{
    public class ProjectSummary
    {
        public string ProjectId { get; set; }
        public DateTime DateFrom { get; set; }
        public DateTime DateTo { get; set; }
        public double TotalHours { get; set; }
        public HashSet<string> Comments { get; set; }
        public List<DateTime> Dates { get; set; }
    }
}
