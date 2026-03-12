using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfCalc
{
    public class Table2Row
    {
        public string OrganizationType { get; set; } = string.Empty;

        public int TotalObjects { get; set; }
        public int EquippedWithSoue { get; set; }
        public int ValidSoueActivations { get; set; }
    }
}
