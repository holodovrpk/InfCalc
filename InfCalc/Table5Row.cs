using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfCalc
{
    public class Table5Row
    {
        public string OrganizationType { get; set; } = string.Empty;

        public int TotalObjects { get; set; }
        public int TrainingObjects { get; set; }
        public int TotalTrainings { get; set; }

        public int ManagersCount { get; set; }
        public double ManagersPercent { get; set; }

        public int WorkersCount { get; set; }
        public double WorkersPercent { get; set; }

        public int StudentsCount { get; set; }
        public double StudentsPercent { get; set; }

        public int SecurityCount { get; set; }
        public double SecurityPercent { get; set; }
    }
}
