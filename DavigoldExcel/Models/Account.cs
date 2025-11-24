using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DavigoldExcel.Models
{
    public class Account
    {
        public int Id { get; set; }

        public string AccountName { get; set; }

        public string AccountNumber { get; set; }

        public bool Selection { get; set; }

        public ComboBoxModel AccountType { get; set; }

        public ComboBoxModel OperatorType { get; set; } = new ComboBoxModel() { Label = "+", Value = "+" };

        public ComboBoxModel Period { get; set; }

    }

    public class AccountGroup
    {

        public string Name { get; set; }

        public bool Selection { get; set; }

        public ComboBoxModel AccountType { get; set; }

        public ComboBoxModel OperatorType { get; set; } = new ComboBoxModel() { Label = "+", Value = "+" };

        public ComboBoxModel Period { get; set; }

    }
}
