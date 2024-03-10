using System;
using System.IO;
using System.ServiceProcess;
using System.Diagnostics;
using System.Management;


namespace PrintSpoolerMonitorService
{
    public partial class PrintSpoolerMonitor : ServiceBase
    {
        private ManagementEventWatcher watcher;

        public PrintSpoolerMonitor()
        {
            ServiceName = "PrintSpoolerMonitor";
        }

        protected override void OnStart(string[] args)
        {
            // Start listening for Print Spooler service events
            StartMonitoringPrintSpooler();
        }

        protected override void OnStop()
        {
            // Cleanup resources
            StopMonitoringPrintSpooler();
        }

        private void StartMonitoringPrintSpooler()
        {
            try
            {
                // Set up WMI query to monitor Print Spooler service events
                WqlEventQuery query = new WqlEventQuery("SELECT * FROM __InstanceModificationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Service' AND TargetInstance.Name = 'Spooler'");

                // Create event watcher
                watcher = new ManagementEventWatcher(query);
                watcher.EventArrived += SpoolerServiceEventArrived;
                watcher.Start();

                LogEvent("Monitoring Print Spooler service...");
            }
            catch (Exception ex)
            {
                LogEvent($"Error starting Print Spooler service monitoring: {ex.Message}");
            }
        }

        private void StopMonitoringPrintSpooler()
        {
            try
            {
                if (watcher != null)
                {
                    watcher.Stop();
                    watcher.Dispose();
                }
            }
            catch (Exception ex)
            {
                LogEvent($"Error stopping Print Spooler service monitoring: {ex.Message}");
            }
        }

        private void SpoolerServiceEventArrived(object sender, EventArrivedEventArgs e)
        {
            try
            {
                // Check if the event is related to a change in the Print Spooler service status
                PropertyData property = e.NewEvent.Properties["TargetInstance"];
                if (property != null)
                {
                    ManagementBaseObject targetInstance = (ManagementBaseObject)property.Value;
                    string serviceName = targetInstance["Name"].ToString();
                    string serviceState = targetInstance["State"].ToString();

                    if (serviceName == "Spooler")
                    {
                        if (serviceState == "Running")
                        {
                            // Print Spooler service has started, create text file and log event
                            CreateTextFile();
                            LogEvent("Print Spooler service has started.");
                        }
                        else
                        {
                            // Print Spooler service has stopped or paused, log event
                            LogEvent($"Print Spooler service state changed to: {serviceState}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogEvent($"Error processing Print Spooler service event: {ex.Message}");
            }
        }

        private void CreateTextFile()
        {
            try
            {
                // Create a text file in C:\temp\
                string filePath = @"C:\temp\PrintSpooler.txt";
                string content = "Print Spooler service is running at " + DateTime.Now.ToString();
                File.WriteAllText(filePath, content);
            }
            catch (Exception ex)
            {
                // Log the exception
                LogEvent($"Error creating text file: {ex.Message}");
            }
        }

        private void LogEvent(string message)
        {
            try
            {
                EventLog.WriteEntry(ServiceName, message, EventLogEntryType.Information);
            }
            catch (Exception ex)
            {
                // Log any exception encountered while writing to the event log
                Console.WriteLine($"Error writing to Event Log: {ex.Message}");
            }
        }
    }
}
