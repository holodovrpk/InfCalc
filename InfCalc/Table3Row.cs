using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InfCalc
{
    public class Table3Row
    {
        public string OrganizationType { get; set; } = string.Empty;

        public int TotalObjects { get; set; }
        public int HasAlgorithms { get; set; }
        public int UpdatedAlgorithms { get; set; }
    }
}
