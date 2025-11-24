using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DavigoldExcel.Models
{

    public class LabelJson
    {
        public string NameEn { get; set; }
        public string NameFr { get; set; }
        public string LabelEn { get; set; }
        public string LabelFr { get; set; }
        public string Slug { get; set; }
    }

    public class Label
    {
        public int Id { get; set; }
        public string ChangeIn { get; set; }
        public string Module { get; set; }
        public string Table { get; set; }
        public string Form { get; set; }
        public bool? IsAddin { get; set; }
        public bool? IsAddinUpload { get; set; }


        public List<LabelJson> Labels { get; set; }
    }

    public class DropdownOption
    {
        public string EnValue { get; set; }
        public string EnLabel { get; set; }
        public string FrValue { get; set; }
        public string FrLabel { get; set; }
        public string Slug { get; set; }
    }

    public class Dropdown
    {
        public int Id { get; set; }
        public int TenantId { get; set; }
        public string DefaultValue { get; set; }
        public string Sort{ get; set; }
        public string dropdown { get; set; }
        public string MainObject { get; set; }
        public List<DropdownOption> Options { get; set; }

    }

    public class Feature
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Slug { get; set; }
        public bool Selection { get; set; }
        public List<Feature> Screens { get; set; }
    }
    

    public class Tenant
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public List<string> ClientType { get; set; }

        public List<Label> Labels { get; set; }

        public List<Dropdown> Dropdowns { get; set; }

        public List<Feature> Features { get; set; }

        public List<Module> Modules { get; set; }

    }
}
