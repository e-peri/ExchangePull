using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;

namespace ExchnagePullController
{
    

    public class PSHelper
    {
        static public PSHelper Helper { get; } = new PSHelper();

        private PSHelper() { }

        public PSCommand CreateCommand(string command, Dictionary<string, string> parameters)
        {
            if (string.IsNullOrWhiteSpace(command))
            {
                Trace.TraceError("CreateCommand - missing command");
                return null;
            }

            var psCommand = new PSCommand();
#if DEBUG
            // TODO - decide how to silnce / activate
            //PrintCommandParams(parameters);
#endif
            psCommand.AddCommand(command);
            if (parameters != null)
            {
                foreach (var param in parameters)
                {
                    if (string.IsNullOrWhiteSpace(param.Value))
                    {
                        // TODO - do we really need it
                        psCommand.AddParameter(param.Key);
                    }
                    else
                    {
                        psCommand.AddParameter(param.Key, param.Value);
                    }
                }
            }

            return psCommand;
        }

        public void PrintCommandParams(Dictionary<string, string> commandParameters)
        {
            if (commandParameters != null && commandParameters.Count > 0)
            {
                Trace.TraceInformation("PrintCommandParams - command parameters:");
                foreach (var param in commandParameters)
                {
                    Trace.TraceInformation("{0} - {1}", param.Key, param.Value);
                }
            }
        }

        public void PrintCollectionProperties(Collection<PSObject> objects)
        {
            foreach (var item in objects)
            {
                PrintObjectProperties(item);
            }
        }

        public void PrintObjectProperties(PSObject obj)
        {
            Trace.TraceInformation("PrintProperties - print properties");
            foreach (var prop in obj.Properties)
            {
                // TODO - check if null
                if (prop.Value is Object)
                {
                    Trace.TraceInformation("name: {0} - val: {1}", prop.Name, prop.Value.ToString());
                }
                else
                {
                    Trace.TraceInformation("{0} value is null", prop.Name);
                }
            }
        }

        public string GetCommandMaxResults(int requestedMaxResults)
        {
            string maxResults = requestedMaxResults <= 0 ? "unlimited" : requestedMaxResults.ToString();
            Trace.TraceInformation("GetCommandMaxResults - maxResults: " + maxResults);
            return maxResults;
        }
    }
}
