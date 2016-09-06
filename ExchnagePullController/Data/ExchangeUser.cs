using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExchnagePullController.Data
{
    [Serializable()]
    class ExchangeUser
    {
        private ExchangeUser()
        {
            _Devices = new List<ExchangeDevice>();
            _UserProperties = new Dictionary<string, string>();
        }

        public ExchangeUser(Mailbox mailBox) : this()
        {
            Email = mailBox.PrimaryEmail;
            DisplayName = mailBox.DisplayName;
            //ExchangeIdentity = mailBox.ExchangeIdentity;
        }

        
        public string Email { get; set; }
        public string DisplayName { get; set; }
        //public string ExchangeIdentity { get; }
        //public string Country { get; set; }
        //public string Department { get; set; }

        public string Devices
        {
            get { return _Devices.ToString(); }
        }

        public string Properties
        {
            get { return _UserProperties.ToString(); }
        }

        private Dictionary<string, string> _UserProperties { get; }
        private List<ExchangeDevice> _Devices { get; }
    }
}
