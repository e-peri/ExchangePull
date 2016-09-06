using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

/*
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
*/

//using ExchnagePullController;
//using ExchnagePullController.Data;


namespace ExchangeConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Trace and logs
            // http://olondono.blogspot.com/2008/01/about-trace-listeners.html

            ExchnagePullController.ExchnagePullController controller = new ExchnagePullController.ExchnagePullController();
            
            // TODO - read config from file
            controller.Init("");
            var b = controller.Pull();

            Console.WriteLine("Press any key.....");
            Console.ReadKey();
        }
        
    }
}


// tests
/*
            SecureString pass = new SecureString();
            string decr = DecryptString("RaALaiNlGTiFsV4SC+mRXw=="); //"faWYZEmEQd0pP/QHrTPG7Q=="); //"RaALaiNlGTiFsV4SC+mRXw==";
            foreach (var c in decr.ToCharArray())
            {
                pass.AppendChar(c);
            }


            var pool = PSConnectionPool.ConnectionPool;
            pool.Init(5);
            var t = pool.Connect("https://outlook.office365.com/powershell-liveid/?SerializationLevel=Full", "zvi.agmon@landesk.com", pass, 120);

            int y = pool.Capacity;
            var conn = pool.Acquire();
            if(conn != null)
            if (conn.ConnectRemoteExchange("https://outlook.office365.com/powershell-liveid/?SerializationLevel=Full", "zvi.agmon@landesk.com", pass))
            {
                Console.WriteLine("Coonection is : {0}", conn.IsOpen() ? "Open" : "Close");
                
                conn.Event_DataAdded += OnEventDataAdded;
                conn.SendAsync(CMD_GET_CASMAILBOX, new Dictionary<string, string>
                {
                    {PARAM_FILTER, PARAM_DEVICE_PARTNER},
                    {PARAM_RESULT_SIZE, PSHelper.Helper.GetCommandMaxResults(5)}
                });

                Collection<Mailbox> mailboxQueue = new Collection<Mailbox>();
                while (conn.ResoultAvialoble())
                {
                    var resoult = conn.TakeResoult();
                    Mailbox mb = new Mailbox()
                    {
                        PrimaryEmail = resoult.Members[PROP_PRIMARY_EMAIL].Value as string,
                        DisplayName = resoult.Members[PROP_DISPLAY_NAME].Value as string,
                        //ExchangeIdentity = ""
                    };

                    mailboxQueue.Add(mb);
                }

                // you can start pull when the event is trigred
                while (conn.ResoultAvialoble())
                {
                    PSObject obj = conn.TakeResoult();
                    Trace.TraceInformation(obj.ToString());
                }
                

                //paramsDict.Clear();
                //paramsDict.Add(PARAM_MAILBOX, mailboxQueue[0].PrimaryEmail);
                //Collection<PSObject> ret = conn.Send(CMD_GET_DEVICE_STATISTICS, paramsDict, 100);
                Collection<PSObject> ret = conn.Send(CMD_GET_DEVICE_STATISTICS, new Dictionary<string, string>
                {
                    {PARAM_MAILBOX, mailboxQueue[0].PrimaryEmail}
                }, 100);

                PSHelper.Helper.PrintCollectionProperties(ret);
                

                pool.Release(conn);

                pool.Dispose();
                
            }*/

