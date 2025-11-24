using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DavigoldExcel.Models
{
    public class User
    {
        public int Id { get; set; }
        public string Username { get; set; }
        public string Email { get; set; }
        public string Name { get; set; }
        public bool ShowDownloadButton { get; set; }
        public bool ShowUploadButton { get; set; }
        public bool ShowHideShowFields { get; set; }
        public int? TenantId { get; set; }
        public Tenant Tenant { get; set; }
        public List<Kpi> Kpis { get; set; }
        public List<KeyFigureKpi> KeyFigureKpis { get; set; }
        public List<Account> Accounts { get; set; }
        public List<Chart> Charts { get; set; }
        public List<Share> Shares { get; set; }
        public List<Fund> Funds { get; set; }

        public List<string> AccountGroups { get; set; }
    }
}
