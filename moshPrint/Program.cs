using System;
using System.ServiceProcess;

namespace PrintSpoolerMonitorService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new PrintSpoolerMonitor()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
