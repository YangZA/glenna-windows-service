using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace glenna_service
{
    public partial class GlennaService : ServiceBase
    {
        private int eventId = 1;
        private EventLog eventLog;

        public enum ServiceState
        {
            SERVICE_STOPPED = 0x00000001,
            SERVICE_START_PENDING = 0x00000002,
            SERVICE_STOP_PENDING = 0x00000003,
            SERVICE_RUNNING = 0x00000004,
            SERVICE_CONTINUE_PENDING = 0x00000005,
            SERVICE_PAUSE_PENDING = 0x00000006,
            SERVICE_PAUSED = 0x00000007,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ServiceStatus
        {
            public int dwServiceType;
            public ServiceState dwCurrentState;
            public int dwControlsAccepted;
            public int dwWin32ExitCode;
            public int dwServiceSpecificExitCode;
            public int dwCheckPoint;
            public int dwWaitHint;
        };

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);

        public GlennaService()
        {
            InitializeComponent();
            eventLog = new EventLog();

            if (!EventLog.SourceExists("Glenna"))
            {
                EventLog.CreateEventSource(
                    "Glenna", "GlennaLog");
            }
            eventLog.Source = "Glenna";
            eventLog.Log = "GlennaLog";
        }

        protected override void OnStart(string[] args)
        {
            // Update the service state to Start Pending.
            ServiceStatus serviceStatus = new ServiceStatus
            {
                dwCurrentState = ServiceState.SERVICE_START_PENDING,
                dwWaitHint = 100000
            };
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // Update the service state to Running.
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            eventLog.WriteEntry("Glenna service has started.");

            // Set up a timer that triggers every minute.
            Timer timer = new Timer
            {
                Interval = 60000 // 60 seconds
            };
            timer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            timer.Start();
        }

        protected override void OnStop()
        {
            eventLog.WriteEntry("Glenna Service is stopping.");
        }

        public void OnTimer(object sender, ElapsedEventArgs args)
        {
            CheckDpsMeter();
            // TODO: Upload DPS reports automatically
        }

        private void CheckDpsMeter()
        {
            string url = "https://www.deltaconnected.com/arcdps/x64/d3d9.dll.md5sum";
            eventLog.WriteEntry("Polling arcDPS repository.", EventLogEntryType.Information, eventId++);
            string remoteChecksum = FetchRemoteHash(url);
            string localChecksum = CalculateMD5(@"C:\Program Files\Guild Wars 2\bin64\d3d9.dll");

            if (remoteChecksum == localChecksum)
            {
                eventLog.WriteEntry("Installed arcDPS is the most recent version", EventLogEntryType.Information, eventId++);
            }
            else
            {
                InstallArcDps(@"C:\Program Files\Guild Wars 2\bin64\d3d9.dll");
            }
        }

        private void InstallArcDps(string fileName)
        {
            eventLog.WriteEntry("Fetching most recent version of arcDPS", EventLogEntryType.Information, eventId++);

            WebClient webClient = new WebClient();
            try
            {
                File.Delete(fileName);
                webClient.DownloadFile("https://www.deltaconnected.com/arcdps/x64/d3d9.dll", fileName);
            }
            catch (Exception e)
            {
                eventLog.WriteEntry(e.Message, EventLogEntryType.Error, eventId++);
            }

            eventLog.WriteEntry("Installed the latest build of arcDPS", EventLogEntryType.Information, eventId++);
        }

        private string FetchRemoteHash(string url)
        {
            WebClient webClient = new WebClient();
            try
            {
                string hash = webClient.DownloadString(url);
                hash = hash.Replace("  x64/d3d9.dll","").Trim();
                return hash;
            }
            catch (Exception e)
            {
                eventLog.WriteEntry(e.Message, EventLogEntryType.Error, eventId++);
            }
            return null;
        }


        static string CalculateMD5(string filename)
        {
            using (var md5 = MD5.Create())
            {
                using (var stream = File.OpenRead(filename))
                {
                    var hash = md5.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                }
            }
        }
    }
}
