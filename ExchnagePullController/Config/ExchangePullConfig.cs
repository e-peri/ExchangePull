using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace ExchnagePullController
{
    [Serializable()]
    public class ExchangeServerConfig
    {
        private const string OFFICE_365_POWERSHELL_PART = "powershell-liveid";
        private const string EXCHANGE_POWERSHELL_PART = "PowerShell";
        private const string EXCHANGE_POWERSHELL_CONNECT_PATTERN = "https://{0}/{1}/?SerializationLevel=Full";
        

        [XmlIgnore]
        internal string UserName { get; set; }

        [XmlIgnore]
        internal SecureString Password { get; set; }

        [XmlElement("ServerURI")]
        public string ServerURI { get; set; } = "outlook.office365.com";

        [XmlIgnore]
        internal string ConnectionString
        {
            get
            {
                if (string.IsNullOrEmpty(ServerURI))
                {
                    return "";
                }
                string serverName = ServerURI;
                string connectionString = string.Format(EXCHANGE_POWERSHELL_CONNECT_PATTERN, serverName, (serverName.Contains("office365") ? OFFICE_365_POWERSHELL_PART : EXCHANGE_POWERSHELL_PART));
                return connectionString;
            }
        }
    }

    [Serializable()]
    public class PushServerConfig
    {
        [XmlIgnore]
        public string UserName { get; set; }

        [XmlIgnore]
        public SecureString SecurePassword { get; set; }

        [XmlElement("ServerURI")]
        public string ServerURI { get; set; } = "";
    }

    [Serializable()]
    public class ExchangePullConfig
    {
        // single instance
        static public ExchangePullConfig Configuration { get; } = new ExchangePullConfig();

        private ExchangePullConfig() { }

        [XmlElement("MailboxMaxResults")]
        public int MailboxMaxResults { get; set; } = 10; // -1 = all 

        [XmlElement("ScanTimeOutMin")]
        public int ScanTimeOutMin { get; set; } = 15;

        [XmlElement("MaxConnections")]
        public uint MaxConnections { get; set; } = 1;

        [XmlElement("LastSyncInDays")]
        public int LastSyncInDays { get; set; } = 7;

        [XmlElement("ConnectRetryNum")]
        public int ConnectRetryNum { get; set; } = 3;

        [XmlElement("ConnectRetryTimeout")]
        public int ConnectRetryTimeout { get; set; } = 120;

        [XmlElement("ExchangeServerConfig")]
        public ExchangeServerConfig Exchangeconfig { get; set; } = new ExchangeServerConfig();

        [XmlElement("PushServerConfig")]
        public PushServerConfig PushConfig { get; set; } = new PushServerConfig();

    }

    /*
    [Serializable()]
    [XmlRoot("PullConfigCollection")]
    public class PullConfigCollection
    {
        [XmlArray("Configs")]
        [XmlArrayItem("ExchangePullConfig", typeof(ExchangePullConfig))]
        public ExchangePullConfig[] Config { get; }      
    }
    */
}
