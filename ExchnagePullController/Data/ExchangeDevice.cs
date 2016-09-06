using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace ExchnagePullController.Data
{
    class ExchangeDevice
    {
        public ExchangeDevice(string deviceId)
        {
            _DeviceProperties = new Dictionary<string, string>();
            ID = deviceId;
        }
        
        public string ID { get; }

        public string DeviceProperties
        {
            get { return _DeviceProperties.ToString();  }
        }


        private Dictionary<string, string> _DeviceProperties { get; set; }
    }
}
