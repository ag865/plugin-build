using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DavigoldExcel.Models
{
    public  class Module
    {
        public string Name { get; set; }
        public string Label { get; set; }
        public List<Module> SubModules { get; set; }
    }
}
