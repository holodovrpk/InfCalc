using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfCalc
{
    public class Table4Row
    {
        public string Municipality { get; set; } = string.Empty;

        public int MunicipalTrained { get; set; }
        public int MunicipalPlannedCurrentYear { get; set; }
        public int MunicipalPlannedNextYear { get; set; }

        public string OrganizationType { get; set; } = string.Empty;

        public int TotalObjects { get; set; }
        public int OrgTrained { get; set; }
        public int OrgPlannedCurrentYear { get; set; }
        public int OrgPlannedNextYear { get; set; }

        public bool IsMunicipalityTotalRow { get; set; }
        public bool IsGrandTotalRow { get; set; }
    }
}
