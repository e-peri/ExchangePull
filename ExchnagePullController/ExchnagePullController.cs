using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExchnagePullController.Data;
using ExchnagePullController.Util;

//From microsoft documentaion - realte to Exchange 2016
//The Get-ActiveSyncDeviceStatistics cmdlet will be removed in a future version of Exchange. Use the Get-MobileDeviceStatistics cmdlet instead. 
//const string CMD_GET_EXCHANGESERVER = "Get-Exchangeserver";
//const string CMD_GET_DEVICE_STATISTICS = "Get-ActiveSyncDeviceStatistics";

namespace ExchnagePullController
{
    // tasks
    // connect
    // pump
    // mailbox pull
    // device pull

    public class ExchnagePullController
    {
        #region Class data

        public enum PullResoult { Error = -1, Success = 0 }

        internal ConcurrentQueue<Task> tasksQueue;
        internal ConcurrentQueue<Mailbox> mailboxQueue;
        internal ConcurrentQueue<ExchangeUser> usersQueue;
        private Object tasksLock;
        
        internal bool mailboxesPullDone;

        #endregion


        public ExchnagePullController()
        {
            tasksQueue = new ConcurrentQueue<Task>();
            mailboxQueue = new ConcurrentQueue<Mailbox>();
            usersQueue = new ConcurrentQueue<ExchangeUser>();
            tasksLock = new Object();
            mailboxesPullDone = false;
        }

        public bool Init(string configFilePath)
        {
            var configCollection = ExchangePullConfig.Configuration;

            SecureString pass = new SecureString();
            string decr = DecryptString("RaALaiNlGTiFsV4SC+mRXw=="); //"faWYZEmEQd0pP/QHrTPG7Q=="); //"RaALaiNlGTiFsV4SC+mRXw==";
            foreach (var c in decr.ToCharArray())
            {
                pass.AppendChar(c);
            }
            configCollection.Exchangeconfig.UserName = "zvi.agmon@landesk.com";
            configCollection.Exchangeconfig.Password = pass;

            configCollection.PushConfig.UserName = "zvi.agmon@landesk.com";
            
            return true;
        }

        #region Decrypt - temp
        // for test pepose unly
        static string DecryptString(string inputString)
        {
            MemoryStream memStream = null;
            try
            {
                byte[] key = {};
                byte[] IV = {12, 21, 43, 17, 57, 35, 67, 27};
                string encryptKey = "aXb2uy4z"; // MUST be 8 characters
                key = Encoding.UTF8.GetBytes(encryptKey);
                byte[] byteInput = new byte[inputString.Length];
                byteInput = Convert.FromBase64String(inputString);
                DESCryptoServiceProvider provider = new DESCryptoServiceProvider();
                memStream = new MemoryStream();
                ICryptoTransform transform = provider.CreateDecryptor(key, IV);
                CryptoStream cryptoStream = new CryptoStream(memStream, transform, CryptoStreamMode.Write);
                cryptoStream.Write(byteInput, 0, byteInput.Length);
                cryptoStream.FlushFinalBlock();
                Encoding encoding1 = Encoding.UTF8;
                return encoding1.GetString(memStream.ToArray());
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "";
            }
        }

        #endregion

        // dispose

        public PullResoult Pull()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            //var str1 = SerializeHelper.Instanse.ToJson(usersQueue.ToArray());
            //str1 = SerializeHelper.Instanse.ToXml(usersQueue.ToArray());

            var config = ExchangePullConfig.Configuration;

            var pool = PSConnectionPool.ConnectionPool;
            pool.Init(config.MaxConnections);
            pool.Event_ConnectionReady += PullEvent_ConnectionReady; 

            var exchConfig = config.Exchangeconfig;
            var connected = pool.Connect(exchConfig.ConnectionString, exchConfig.UserName, exchConfig.Password, config.ConnectRetryTimeout);

            Task.WaitAll(tasksQueue.ToArray());

            // NOT working
            // http://stackoverflow.com/questions/8084748/dealing-with-ews-throttling-policies
            // https://blogs.msdn.microsoft.com/mstehle/2010/11/09/ews-best-practices-understand-throttling-policies/
            //var test = connection.Send("Get-ThrottlingPolicy", null, 100);

            // if we use just one connection
            if (pool.Capacity == 1)
            {
                // start - it will be executed after mailboxes is done
                var exchangeUserDataTask = Task.Run(() => PullExchangeUserDataTask());
                exchangeUserDataTask.Wait();
            }

            string file = Directory.GetCurrentDirectory() + "\\resoults.json";
            SerializeHelper.Instanse.WriteJsonFile(file, usersQueue.ToArray());
            

            //var pushDataTask = Task.Run(() => PushData());
            //pushDataTask.Wait();

            //var writeDataToFileTask = Task.Run(() => WriteDataToFile());
            //writeDataToFileTask.Wait();



            pool.Dispose();

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            Trace.TraceInformation("Pull time - {0:hh\\:mm\\:ss}", ts);
            
            return PullResoult.Success;
        }
        

        private void PullEvent_ConnectionReady(object sender, int e)
        {
            lock (tasksLock)
            {
                if (tasksQueue.Count == 0)
                {
                    mailboxesPullDone = false;
                    tasksQueue.Enqueue(Task.Run(() => PullMailboxesTask()));
                }
                else
                {
                    tasksQueue.Enqueue(Task.Run(() => PullExchangeUserDataTask()));
                }
            }
        }

        
        internal bool PullMailboxesTask()
        {
            const string CMD_GET_CASMAILBOX = "Get-CASMailbox";
            const string PARAM_FILTER = "-Filter";
            const string PARAM_DEVICE_PARTNER = "hasactivesyncdevicepartnership -eq $true -and -not displayname -like \"CAS_{*\"";
            const string PARAM_RESULT_SIZE = "-ResultSize";

            // TODO - do i need to protect aginst function reuse (a call while it is running)
            //if (tasksQueue.Count > 0)
            //    return false;

            var pool = PSConnectionPool.ConnectionPool;
            var connection = pool.Acquire();
            if (connection == null)
            {
                Trace.TraceError("PullMailboxesTask - Connection quueue is empty");
            }

            var maxMailboxes = ExchangePullConfig.Configuration.MailboxMaxResults;
            
            if (connection != null)
            {
                var pull = Task.Run(() => MailboxesQueuePump(connection));

                connection.Event_DataAdded += On_Mailboxes_Added;   // not sure it's  needed if we use pump
                connection.Event_DataDone += On_Mailboxes_Done;     // not sure it's  needed

                connection.SendAsync(CMD_GET_CASMAILBOX, new Dictionary<string, string>
                {
                    {PARAM_FILTER, PARAM_DEVICE_PARTNER},
                    {PARAM_RESULT_SIZE, PSHelper.Helper.GetCommandMaxResults(maxMailboxes)}
                });

                connection.Event_DataAdded -= On_Mailboxes_Added;
                connection.Event_DataDone -= On_Mailboxes_Done;

                pull.Wait();
                
                pool.Release(connection);
                connection = null;
                
                // TODO - test it
                // we have a free connection - try to add another pull task
                //tasksQueue.Enqueue(Task.Run(() => PullExchangeUserDataTask()));
            }
            else
            {
                // should not happne - we are trigred by connection avilabole 
                Trace.TraceError("Error no connction avilabole for Mailboxes");
            }

            return true;
        }

      
        private void On_Mailboxes_Done(object sender, int e)
        {
            Trace.TraceInformation("Finised mailbox pull");
            mailboxesPullDone = true; // end the mailbox pump
        }

        private void On_Mailboxes_Added(object sender, int e)
        {
            // TODO - we can do item to mailbox here here, but it might slow the pull process. to check
            // 
        }

        internal bool MailboxesQueuePump(PSExchangeConnection connection)
        {
            const string PROP_PRIMARY_EMAIL = "PrimarySmtpAddress";
            const string PROP_DISPLAY_NAME = "DisplayName";
            
            while (!mailboxesPullDone)
            {
                // TODO - use Take and resultsCollection and avoid the sleep
                while (connection.ResoultAvialoble())
                {
                    var resoult = connection.TakeResoult();
                    if (resoult != null && resoult.GetType() == typeof(PSObject))
                    {
                        Mailbox mb = new Mailbox()
                        {
                            PrimaryEmail = resoult.Members[PROP_PRIMARY_EMAIL].Value as string,
                            DisplayName = resoult.Members[PROP_DISPLAY_NAME].Value as string,
                            //ExchangeIdentity = ""
                        };

                        Trace.TraceInformation("Mailbox {0}: {1} - {2}", mailboxQueue.Count, mb.DisplayName,
                            mb.PrimaryEmail);

                        mailboxQueue.Enqueue(mb);
                    }
                }

                Thread.Sleep(200);
            }
            
            return true;
        }


        internal bool PullExchangeUserDataTask()
        {
            const string CMD_GET_DEVICE_STATISTICS = "Get-ActiveSyncDeviceStatistics";
            const string PARAM_MAILBOX = "-Mailbox";
            const string PROP_LAST_SUCCESS_SYNC = "LastSuccessSync";

            const string CMD_GET_MAILBOX = "Get-Mailbox";
            const string PARAM_IDENTITY = "-Identity";
            const string PROP_USAGE_LOCATION = "UsageLocation";

            const string PROP_DEPARTMENT = "Department";
            const string CMD_GET_USER = "Get-User";

            int count = 0;
            var pool = PSConnectionPool.ConnectionPool;
            var connection = pool.Acquire();
            if (connection == null)
            {
                Trace.TraceError("PullExchangeUserData - Connection quueue is empty");
            }

            while (!mailboxesPullDone || !mailboxQueue.IsEmpty)
            {
                Mailbox mailbox = null;
                var resoult = mailboxQueue.TryDequeue(out mailbox);
                if (resoult && mailbox != null)
                {
                    var devices = connection.Send(CMD_GET_DEVICE_STATISTICS, new Dictionary<string, string>
                        {
                            {PARAM_MAILBOX, mailbox.PrimaryEmail}
                        }, 100);
                    
                    TimeSpan ta = Convert.ToDateTime(devices[0].Members[PROP_LAST_SUCCESS_SYNC].Value).Subtract( DateTime.Now);

                    var location = connection.Send(CMD_GET_MAILBOX, new Dictionary<string, string>
                    {
                        { PARAM_IDENTITY, mailbox.PrimaryEmail}
                    },100);
                    var usageLocationObj = location.Count > 0 ? location[0].Members[PROP_USAGE_LOCATION].Value : "";

                    //OrgPersonPresentationObject

                    var org = connection.Send(CMD_GET_USER, new Dictionary<string, string>
                    {
                        {PARAM_IDENTITY, mailbox.PrimaryEmail }
                    }, 100);
                    var departmentObj = org.Count > 0 ? org[0].Members[PROP_DEPARTMENT].Value : "";

                    var user = new ExchangeUser(mailbox);
                    usersQueue.Enqueue(user);

                    count++;

                    Trace.TraceInformation("User {0} has {1} devices", mailbox.PrimaryEmail, devices.Count);
                    Trace.TraceInformation("In {0} - Org {1} ", usageLocationObj, departmentObj);
                }
            }

            Trace.TraceInformation("Connection #{0} finished - pull {1} mailboxes", connection.ConnectionId, count);
            
            pool.Release(connection);
            
            return true;
        }



    }
}
