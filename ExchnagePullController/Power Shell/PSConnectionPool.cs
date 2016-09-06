using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExchnagePullController
{
    public class PSConnectionPool : IDisposable
    {
        #region Class data

        private ConcurrentQueue<PSExchangeConnection> connectionPull;
        private uint maxElemnts;
        private int remainingRetries; 

        public event EventHandler<int> Event_ConnectionReady;

        #endregion

        public static PSConnectionPool ConnectionPool { get; } = new PSConnectionPool();

        #region Init and Dispose

        private PSConnectionPool()
        {
            maxElemnts = 1;
            remainingRetries = 0;
        }

        public void Init(uint capacity)
        {
            maxElemnts = capacity;
            connectionPull = new ConcurrentQueue<PSExchangeConnection>();
        }

        private void Close()
        {
            if(connectionPull == null)
                return;
            
            foreach (var connection in connectionPull)
            {
                connection.Dispose();
            }
            connectionPull = null;
            maxElemnts = 0;
        }

        public void Dispose()
        {
            Trace.TraceInformation("Diosposong conection pool");

            Event_ConnectionReady = null;

            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (!disposing)
                return;

            Close();
        }

        #endregion

        /// <summary>
        /// Create and connect maxElemnts objects
        /// </summary>
        /// <param name="connectionStr">URL</param>
        /// <param name="userName">User</param>
        /// <param name="password">Password</param>
        /// <param name="timeInSec">Timeout in secounds</param>
        /// <returns>True if finised on time. False if timed out (number of connection is unknown)</returns>
        public bool Connect(string connectionStr, string userName, SecureString password, int timeInSec)
        {
            if (string.IsNullOrWhiteSpace(connectionStr) || string.IsNullOrWhiteSpace(userName) || password == null)
                return false;

            if (connectionPull == null || connectionPull.Count > 0)
                return false;

            // TODO - impel
            //remainingRetries = ExchangePullConfig.Configuration.ConnectRetryNum;

            var tasks = new Task[maxElemnts];
            for (var i = 0; i < maxElemnts; i++)
            {
                tasks[i] = Task.Factory.StartNew(() =>
                {
                    PSExchangeConnection psConn = PSExchangeConnection.CreateConnection();
                    try
                    {
                        psConn.ConnectRemoteExchange(connectionStr, userName, password);
                        if (psConn.IsOpen())
                        {
                            connectionPull.Enqueue(psConn);
                            Event_ConnectionReady?.Invoke(this, connectionPull.Count);
                        }
                    }
                    catch (Exception e)
                    {
                        Trace.TraceError("InitConnections - exception while connect to exchange: {0}", e.Message);
                    }                    
                });

                // TODO - check if needed
                Thread.Sleep(1000);
            }

            bool inTime = Task.WaitAll(tasks, timeInSec * 1000);

            // TODO - add retry

            return inTime;
        }

        public int Capacity { get { return connectionPull.Count; } }

        /// <summary>
        /// get a connection for ongoinf use. you must release when done
        /// </summary>
        /// <returns></returns>
        public PSExchangeConnection Acquire()
        {
            // TODO - do i need to check if it is busy?
            PSExchangeConnection ret = null;
            if (connectionPull.Count > 0)
                connectionPull.TryDequeue(out ret);
            
            return ret;
        }

        /// <summary>
        /// "relase" the resource. do not use the object after calling
        /// </summary>
        /// <param name="connection"></param>
        public void Release(PSExchangeConnection connection)
        {
            connectionPull.Enqueue(connection);
        }
    }
}
