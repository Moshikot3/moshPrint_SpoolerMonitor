using System;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.ServiceProcess;

namespace PrintSpoolerMonitorService
{
    public partial class PrintSpoolerMonitor : ServiceBase
    {
        // Import the necessary Win32 functions
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern bool CreateProcessAsUser(
            IntPtr hToken,
            string lpApplicationName,
            string lpCommandLine,
            IntPtr lpProcessAttributes,
            IntPtr lpThreadAttributes,
            bool bInheritHandles,
            uint dwCreationFlags,
            IntPtr lpEnvironment,
            string lpCurrentDirectory,
            ref STARTUPINFO lpStartupInfo,
            out PROCESS_INFORMATION lpProcessInformation);

        [StructLayout(LayoutKind.Sequential)]
        public struct PROCESS_INFORMATION
        {
            public IntPtr hProcess;
            public IntPtr hThread;
            public uint dwProcessId;
            public uint dwThreadId;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct STARTUPINFO
        {
            public uint cb;
            public string lpReserved;
            public string lpDesktop;
            public string lpTitle;
            public uint dwX;
            public uint dwY;
            public uint dwXSize;
            public uint dwYSize;
            public uint dwXCountChars;
            public uint dwYCountChars;
            public uint dwFillAttribute;
            public uint dwFlags;
            public short wShowWindow;
            public short cbReserved2;
            public IntPtr lpReserved2;
            public IntPtr hStdInput;
            public IntPtr hStdOutput;
            public IntPtr hStdError;
        }

        public PrintSpoolerMonitor()
        {
            ServiceName = "PrintSpoolerMonitor";
        }

        protected override void OnStart(string[] args)
        {
            // Check if a session is active
            bool isSessionActive = IsSessionActive();

            // Start listening for Print Spooler service events if a session is active
            if (isSessionActive)
            {
                LogEvent("Session is active. Starting monitoring Print Spooler service...");
                StartMonitoringPrintSpooler();
            }
            else
            {
                LogEvent("Session is not active. Monitoring of Print Spooler service will not start.");
            }
        }

        protected override void OnStop()
        {
            // Cleanup resources
        }

        private void StartMonitoringPrintSpooler()
        {
            try
            {
                // Set up WMI query to monitor Print Spooler service events
                ManagementEventWatcher watcher = new ManagementEventWatcher("SELECT * FROM __InstanceModificationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Service' AND TargetInstance.Name = 'Spooler'");
                watcher.EventArrived += SpoolerServiceEventArrived;
                watcher.Start();

                LogEvent("Monitoring Print Spooler service...");
            }
            catch (Exception ex)
            {
                LogEvent($"Error starting Print Spooler service monitoring: {ex.Message}");
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
                            // Print Spooler service has started, run script with user permissions
                            RunScriptWithUserPermissions();
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

        private void RunScriptWithUserPermissions()
        {
            try
            {
                string scriptPath = @"C:\temp\test.bat";
                IntPtr hToken = IntPtr.Zero;
                IntPtr hProcess = IntPtr.Zero;
                IntPtr hThread = IntPtr.Zero;

                // Get the token of the active user session
                bool success = WTSQueryUserToken(WTSGetActiveConsoleSessionId(), out hToken);
                if (!success)
                {
                    throw new Exception("Failed to get user token for the active session.");
                }

                // Fill the startup info structure
                STARTUPINFO startupInfo = new STARTUPINFO();
                startupInfo.cb = (uint)Marshal.SizeOf(startupInfo);

                // Create the process as the active user
                PROCESS_INFORMATION processInfo;
                success = CreateProcessAsUser(
                    hToken,
                    null,
                    scriptPath,
                    IntPtr.Zero,
                    IntPtr.Zero,
                    false,
                    0,
                    IntPtr.Zero,
                    null,
                    ref startupInfo,
                    out processInfo);

                if (!success)
                {
                    throw new Exception($"Failed to create process as the active user. Error code: {Marshal.GetLastWin32Error()}");
                }

                // Wait for the process to exit if needed
                // WaitForSingleObject(processInfo.hProcess, INFINITE);

                // Close the handles to avoid resource leaks
                CloseHandle(processInfo.hProcess);
                CloseHandle(processInfo.hThread);
            }
            catch (Exception ex)
            {
                LogEvent($"Error running script with user permissions: {ex.Message}");
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

        private bool IsSessionActive()
        {
            Process[] processes = Process.GetProcesses();
            foreach (Process process in processes)
            {
                try
                {
                    // Check if the process has a valid session ID and belongs to a user
                    if (process.SessionId > 0 && process.MachineName == "." && process.ProcessName != "Idle" && process.ProcessName != "System")
                    {
                        return true;
                    }
                }
                catch (Exception)
                {
                    // Ignore any exceptions thrown while accessing process properties
                }
            }
            return false;
        }

        // Define the necessary Win32 functions
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool CloseHandle(IntPtr hObject);

        [DllImport("Wtsapi32.dll", SetLastError = true)]
        static extern bool WTSQueryUserToken(uint sessionId, out IntPtr phToken);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern uint WTSGetActiveConsoleSessionId();
    }
}
