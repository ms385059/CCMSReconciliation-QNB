using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;

namespace CCMSReconciler
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            bool debugMode = ConfigurationManager.AppSettings["DebugMode"] == "1";

            if (debugMode)
            {
                Service1 service = new Service1();
                service.OnDebug();
                Thread.Sleep(System.Threading.Timeout.Infinite);
            }
            else
            {
                ServiceBase[] ServicesToRun =
                {
                    new Service1()
                };

                ServiceBase.Run(ServicesToRun);
            }
        }
    }
}
