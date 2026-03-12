using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfCalc
{
    public class Table1Row
    {
        public string OrganizationType { get; set; } = string.Empty;

        public int TotalObjects { get; set; }
        public int EquippedWithTs { get; set; }
        public int ValidCalls { get; set; }
        public int PreventedOffenses { get; set; }
        public int DetainedPersons { get; set; }
    }
}
