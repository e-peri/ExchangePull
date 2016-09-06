using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Threading;



namespace ExchnagePullController
{

   public class PSExchangeConnection : IDisposable
    {
        #region Class data

        const string EXCHANGE_CONFIGURATION_NAME = "Microsoft.Exchange";
        const string EXCHANGE_SCHEMA_URI = "http://schemas.microsoft.com/powershell/Microsoft.Exchange";

        // used for object ID (running number)
        private static int count = 0;
        
        private PowerShell powerShell = null;
        private Runspace runSpace = null;
        private WSManConnectionInfo connectionInfo = null;

        private Collection<PSObject> resultsCollection = null;
        private BlockingCollection<PSObject> resoultQueue = null; // maybe use ConcurrentQueue

        public event EventHandler<int> Event_DataAdded;
        public event EventHandler<int> Event_DataDone;

        #endregion
        

        /// <summary>
        /// Creation method. CTor is priavte
        /// </summary>
        /// <returns></returns>
        public static PSExchangeConnection CreateConnection()
        {
            return new PSExchangeConnection();
        }

        #region Init

        private PSExchangeConnection()
        {
            // used to return the collection in sync mode. used as a buffer in async mode
            resultsCollection = new Collection<PSObject>();
            // not in use in sync mode. holds the data in async mode - use Take to get data
            resoultQueue = new BlockingCollection<PSObject>();
            
        }
        
        /// <summary>
        /// Object cleanup
        /// </summary>
        private void Close()
        {
            if (runSpace != null)
            {
                try
                {
                    Trace.TraceInformation("Diosposong Runspace {0} object", runSpace.InstanceId);
                    runSpace.Close();
                    runSpace.Dispose();
                    runSpace = null;
                }
                catch (Exception e)
                {
                    Trace.TraceInformation(e.Message);
                }
            }

            if (powerShell != null)
            {
                try
                {
                    Trace.TraceInformation("Diosposong PowerShell {0} object", powerShell.InstanceId);
                    ClearPowerShell();
                    powerShell.InvocationStateChanged -= Powershell_InvocationStateChanged;
                    powerShell.Dispose();
                    powerShell = null;
                }
                catch (Exception e)
                {
                    Trace.TraceInformation(e.Message);
                }
            }
        }

        /// <summary>
        /// Clear the stream before sending new comman
        /// </summary>
        private void ClearPowerShell()
        {
            if (powerShell == null)
                return;

            powerShell.Streams.ClearStreams();
            powerShell.Commands.Clear();
        }

        private WSManConnectionInfo CreateConnectionInfo(string connectionStr, string schemaUri, string userName, SecureString password)
        {
            try
            {
                return new WSManConnectionInfo(new Uri(connectionStr), schemaUri, new PSCredential(userName, password));
            }
            catch (Exception e)
            {
                Trace.TraceInformation("createConnectionInfo - failed to create: {0}", e.Message);
                throw;
            }
        }

        public void Dispose()
        {
            Trace.TraceInformation("Diosposong conection #{0}", ConnectionId);

            resultsCollection.Clear();
            resoultQueue.Dispose();

            Dispose(true);
            GC.SuppressFinalize(this);

            count--;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing)
                return;

            Close();
        }

        #endregion

        #region Connection Status

        // TODO - check why it crash in the end
        public bool IsOpen()
        {
            if (runSpace == null || powerShell == null)
            {
                return false;
            }

            return (runSpace.RunspaceAvailability != RunspaceAvailability.None) && (runSpace.RunspaceStateInfo.State == RunspaceState.Opened);
        }

        public bool IsBusy()
        {
            return runSpace.RunspaceAvailability == RunspaceAvailability.Busy; 
        }

        public bool IsAvailable()
        {
            return runSpace.RunspaceAvailability == RunspaceAvailability.Available;
        }

        #endregion

        #region Data Access

        public bool ResoultAvialoble()
        {
            return resoultQueue.Count > 0;
        }

        public PSObject TakeResoult()
        {
            return resoultQueue.Take();
        }

        public int ConnectionId { get; internal set; }

        #endregion

        public bool ConnectRemoteExchange(string connectionStr, string userName, SecureString password)
        {
            if (string.IsNullOrWhiteSpace(connectionStr) || string.IsNullOrWhiteSpace(userName) || password.Length == 0)
            {
                Trace.TraceError("Error - parmeters error");
                return false;
            }

            if (IsOpen())
            {
                Trace.TraceError("Error - already open");
                return false;
            }

            Trace.TraceInformation("ConnectRemoteExchange - Start");

            connectionInfo = CreateConnectionInfo(connectionStr, EXCHANGE_SCHEMA_URI, userName, password);
            if (connectionInfo != null)
            {
                // TODO - read from a file and use
                //connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
                //connectionInfo.SkipCACheck = true;
                //connectionInfo.SkipCNCheck = true;
                //connectionInfo.MaximumConnectionRedirectionCount = 4;
                ConnectRemote(EXCHANGE_CONFIGURATION_NAME);
            }
            else
            {
                // TODO - error
                Trace.TraceError("Fail to ConnectRemoteExchange");

                return false;
            }

            count++;
            ConnectionId = count;

            Trace.TraceInformation("ConnectRemoteExchange {0} - Start", ConnectionId);
            
            return IsOpen();
        }



        public Collection<PSObject> Send(string commandText, Dictionary<string, string> parameters, int timeoutSec)
        {
            Trace.TraceInformation("Connection: {0} Send command: {1}", ConnectionId, commandText);

            ClearPowerShell();

            powerShell.Commands = PSHelper.Helper.CreateCommand(commandText, parameters); 
            
            return RunSync(timeoutSec);
        }

        public void SendAsync(string commandText, Dictionary<string, string> parameters)
        {
            Trace.TraceInformation("Connection: {0} SendAsync command: {1}", ConnectionId, commandText);
            
            ClearPowerShell();
            
            powerShell.Commands = PSHelper.Helper.CreateCommand(commandText, parameters);

            RunAsync();
        }

        

        private Collection<PSObject> RunSync(int timeoutSec)
        {
            resultsCollection.Clear();

            bool inTime = Task.Run(() =>
            {
                resultsCollection = powerShell.Invoke();
            }).Wait(timeoutSec * 1000);
            
            if (!inTime)
            {
                Trace.TraceWarning("RunSync timedout");
            }
            
            return resultsCollection;
        }

        /// <summary>
        /// Run async and wait for the operation to finish
        /// the reosults are handled in Output_DataAdded
        /// </summary> 
        private void RunAsync()
        {
            resultsCollection.Clear();

            PSDataCollection<PSObject> asyncOutput = new PSDataCollection<PSObject>();
            asyncOutput.DataAdded += Output_DataAdded;

            // start async calls
            IAsyncResult asyncResult = powerShell.BeginInvoke<PSObject, PSObject>(null, asyncOutput);
            // wait 
            powerShell.EndInvoke(asyncResult);

            // notify - done
            Event_DataDone?.Invoke(this, resoultQueue.Count);

            asyncOutput.DataAdded -= Output_DataAdded;
        }

        private void Output_DataAdded(object sender, DataAddedEventArgs e)
        {
            PSDataCollection<PSObject> myp = (PSDataCollection<PSObject>)sender;

            Collection<PSObject> results = myp.ReadAll();

            foreach (var res in results)
            {
                resoultQueue.Add(res);
                // notify - added
                Event_DataAdded?.Invoke(this, resoultQueue.Count);
            }
        }

        // TODO - not sure we need it
        /// <summary>
        /// called during the connect process
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Powershell_InvocationStateChanged(object sender, PSInvocationStateChangedEventArgs e)
        {
            var psInfo = (PowerShell)sender;

            if (e.InvocationStateInfo.State == PSInvocationState.Running)
            {
                //Trace.TraceInformation("Powershell_InvocationStateChanged id {0} - {1} Running", connectionId, psInfo.Commands);
                return;
            }

            if (e.InvocationStateInfo.State == PSInvocationState.Completed)
            {
                //Trace.TraceInformation("Powershell_InvocationStateChanged {0} - {1} Completed", ConnectionId, psInfo.Commands.Commands);
            }
            else if (e.InvocationStateInfo.State == PSInvocationState.Stopping)
            {
                Trace.TraceInformation("Powershell_InvocationStateChanged - stopping");
            }
            else if (e.InvocationStateInfo.State == PSInvocationState.Stopped)
            {
                Trace.TraceInformation("Powershell_InvocationStateChanged - stopped");
            }
            else if (e.InvocationStateInfo.State == PSInvocationState.Failed)
            {
                string reasonStr = e.InvocationStateInfo.Reason == null ? "" : e.InvocationStateInfo.Reason.ToString();
                Trace.TraceInformation("Powershell_InvocationStateChanged - error reason: " + reasonStr);
                foreach (var errorRecord in powerShell.Streams.Error)
                {
                    Trace.TraceInformation("Powershell_InvocationStateChanged - errorRecord: " + errorRecord.ErrorDetails.Message);
                }

                psInfo.Stop();

                Trace.TraceInformation("Powershell_InvocationStateChanged - command failed, command: " + powerShell.Commands);
                Trace.TraceInformation(powerShell.ToString());
            }
            else
            {
                Trace.TraceInformation("Powershell_InvocationStateChanged - unexpected invocation state: {0}", e.InvocationStateInfo.State.ToString());
            }
        }
        

        private void ConnectLocal()
        {
            throw new NotImplementedException("ConnectLocal");
        }

        private void ConnectRemote(string configurationName)
        {
            Close();
            
            try
            {
                powerShell = PowerShell.Create();
                powerShell.InvocationStateChanged += Powershell_InvocationStateChanged;

                Connect(configurationName);
            }
            catch (Exception e)
            {
                Trace.TraceInformation("ConnectRemote - failed to connect: {0}", e.Message);
                Close();

                // TODO - handle
                //throw;
            }
        }


        private void Connect(string configurationName)
        {
            Trace.TraceInformation("Start Connect uri: {0}", connectionInfo.ConnectionUri);

            InitialSessionState session = InitialSessionState.CreateDefault();
            session.ImportPSModule(new string[] { "MSOnline" });

            runSpace = RunspaceFactory.CreateRunspace(session);
            runSpace.Open();

            // TODO - add check runSpace

            // add exchange remote connection
            var command = new PSCommand();
            command.AddCommand("New-PSSession");
            command.AddParameter("ConfigurationName", configurationName);
            command.AddParameter("ConnectionUri", connectionInfo.ConnectionUri);
            command.AddParameter("Credential", connectionInfo.Credential);
            command.AddParameter("AllowRedirection");
            command.AddParameter("Authentication", "Basic");
            powerShell.Commands = command;
            powerShell.Runspace = runSpace;

            var result = powerShell.Invoke();
            if (result == null || result.Count == 0)
            {
                CheckPowerShellInvokError();

                throw new Exception("failed to connect to exchange server: " + connectionInfo.ConnectionUri);
            }

            // set execution policy Unrestricted
            command = new PSCommand();
            command.AddCommand("Set-ExecutionPolicy");
            command.AddParameter("Scope", "Process");
            command.AddParameter("ExecutionPolicy", "Unrestricted");
            powerShell.Commands = command;
            powerShell.Runspace = runSpace;
            powerShell.Invoke();

            if(CheckPowerShellInvokError())
                return;
            
            // import the session
            object psSessionConnection = result[0];
            PSCommand setVar = new PSCommand();
            setVar.AddCommand("Set-Variable");
            setVar.AddParameter("Name", "ra");
            setVar.AddParameter("Value", result[0]);
            powerShell.Commands = setVar;
            powerShell.Runspace = runSpace;
            powerShell.Invoke();

            if(CheckPowerShellInvokError())
                return;

            command = new PSCommand();
            command.AddScript("Import-PSSession -AllowClobber -Session $ra");
            powerShell.Commands = command;
            powerShell.Runspace = runSpace;
            powerShell.Invoke();

            CheckPowerShellInvokError();
        }

        #region Helpers

        /// <summary>
        /// check error stream and sump it if errors are found
        /// </summary>
        /// <returns>return true if error was found</returns>
        private bool CheckPowerShellInvokError()
        {
            var errors = powerShell.Streams.Error;
            foreach (ErrorRecord current in errors)
            {
                Trace.TraceError("------");
                Trace.TraceError("Error: {0}", current.ErrorDetails.Message);
                //Trace.TraceError("Action: {0}", current.ErrorDetails.RecommendedAction);
                Trace.TraceError("Exception: {0}", current.Exception );
                if(current.Exception.InnerException != null)
                    Trace.TraceError("InnerException: {0}", current.Exception.InnerException);
            }

            return errors.Count > 0;
        }

        #endregion
    }
}
