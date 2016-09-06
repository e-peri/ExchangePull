using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExchnagePullController.Util
{
    public static class TraceHelper
    {
        static TraceHelper()
        {
            // init here

            // Trace and logs
            // http://olondono.blogspot.com/2008/01/about-trace-listeners.html
            // http://www.daveoncsharp.com/2009/09/create-a-logger-using-the-trace-listener-in-csharp/
            // http://stackoverflow.com/questions/576185/logging-best-practices
            // http://stackoverflow.com/questions/16743804/implementing-a-log-viewer-with-wpf
        }

        public static void TraceInformation(string message, params object[] args)
        {
            Trace.TraceInformation(message, args);
        }

        /// <summary>
        /// add time proc etc'
        /// </summary>
        /// <param name="message"></param>
        /// <returns></returns>
        static string FormatMessage(string message)
        {
            return message;
        }
    }
}
