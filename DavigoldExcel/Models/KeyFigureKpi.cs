using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DavigoldExcel.Models
{
    public class KeyFigureKpi
    {
        public int Id { get; set; }

        public string Name { get; set; }
        
        public string KpiType { get; set; }

        public string TableName { get; set; }
        public int TableId{ get; set; }

    }
}
