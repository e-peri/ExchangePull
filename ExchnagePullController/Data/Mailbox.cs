using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExchnagePullController.Data
{
    public class Mailbox
    {
        public Mailbox() { }

        public string PrimaryEmail { get; set; }
        public string DisplayName { get; set; }
        //public string ExchangeIdentity { get; set; }
    }
}
