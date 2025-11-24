using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DavigoldExcel.Models
{
    public class Kpi
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public string Description { get; set; }

        public string Module { get; set; }

        public string SubModule { get; set; }
        public string Suffix { get; set; }
        public int? NoOfDecimals { get; set; }
    }
}
