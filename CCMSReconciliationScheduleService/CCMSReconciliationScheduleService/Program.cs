using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ReconciliationScheduleService
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
                ReconciliationScheduleService service = new ReconciliationScheduleService();
                service.OnDebug();
                Thread.Sleep(System.Threading.Timeout.Infinite);
            }
            else
            {
                ServiceBase[] ServicesToRun =
                {
                    new ReconciliationScheduleService()
                };

                ServiceBase.Run(ServicesToRun);
            }
        }
    }
}
