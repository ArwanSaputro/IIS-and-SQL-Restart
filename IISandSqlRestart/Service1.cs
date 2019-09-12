using APP_COMMON;
using APP_MODEL.ModelData;
using Microsoft.Web.Administration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Timers;

namespace IISandSqlRestart
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }
        protected System.Timers.Timer timer1 = new System.Timers.Timer();
        protected int timeInterval = IISandSqlRestart.Properties.Settings.Default.TimeInterval;
        protected TimeSpan timeAlert = IISandSqlRestart.Properties.Settings.Default.TimeAlert;
        protected TimeSpan timeBusyChecking = IISandSqlRestart.Properties.Settings.Default.TimeBusyChecking;

        protected override void OnStart(string[] args)
        {
            timer1.Elapsed += timer1_Elapsed;
            timer1.Interval = timeInterval;
            timer1.Enabled = true;
        }

        private void timer1_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                ModelEntities db = new ModelEntities();

               
                if (DateTime.Now.Hour == timeAlert.Hours && DateTime.Now.Minute == timeAlert.Minutes && DateTime.Now.Second == timeAlert.Seconds)
                {
                    string Modulname = null;
                    using (var currentProc = Process.GetCurrentProcess())
                    {
                        Modulname = currentProc.MainModule.ModuleName;
                    }
                    var GetLastLogin= db.tbl_Log_System.OrderByDescending(s=>s.Access_DateTime).FirstOrDefault();
                    if(GetLastLogin!=null)
                    {
                        //open IIS Manager
                        var compareTime = GetLastLogin.Access_DateTime.Value.Add(timeBusyChecking);
                        if (DateTime.Now  >= compareTime)
                        {
                            //foreach (var x in new ServerManager().ApplicationPools)
                            //    x.Recycle();
                           var IsIisInstalled = ServiceController.GetServices().Any(s => s.ServiceName.Equals("w3svc", StringComparison.InvariantCultureIgnoreCase));
                            if (IsIisInstalled)
                            {
                                ServerManager serverManager = new ServerManager();
                                ApplicationPoolCollection applicationPoolCollection = serverManager.ApplicationPools;
                                foreach (ApplicationPool applicationPool in applicationPoolCollection)
                                {

                                    // If the applicationPool is stopped, restart it.
                                    if (applicationPool.State == ObjectState.Stopped)
                                    {
                                        applicationPool.Start();
                                        WriteEventLogEntry("IIS Pool " + applicationPool.Name + " has been Started", Modulname);

                                    }
                                    if (applicationPool.State == ObjectState.Started)
                                    {
                                        applicationPool.Stop();
                                        WriteEventLogEntry("IIS Pool " + applicationPool.Name + " has been Stopped", Modulname);
                                        while (true)
                                        {
                                            if (applicationPool.State == ObjectState.Stopped)
                                            {
                                                applicationPool.Start();
                                                WriteEventLogEntry("IIS Pool " + applicationPool.Name + " has been Started", Modulname);
                                                break;
                                            }

                                        }
                                    }

                                }

                                // CommitChanges to persist the changes to the ApplicationHost.config.
                                serverManager.CommitChanges();
                            }
                            ServiceController[] services = ServiceController.GetServices();
                            var service = services.FirstOrDefault(s => s.ServiceName == "MSSQLSERVER");
                            if (service != null)
                            {
                                string strService = "MSSQLSERVER";
                                ServiceController serv = new ServiceController(strService);
                                if (serv != null)
                                {
                                    serv.Stop();
                                    serv.WaitForStatus(ServiceControllerStatus.Stopped);
                                    WriteEventLogEntry("Services " + serv.ServiceName + " has been Stopped", Modulname);

                                    serv.Start();
                                    serv.WaitForStatus(ServiceControllerStatus.Running);
                                    WriteEventLogEntry("Services " + serv.ServiceName + " has been Started", Modulname);
                                }
                            }
                        }

                        //Open cmd
                        //var compareTime = GetLastLogin.Access_DateTime.Value.Add(timeBusyChecking);
                        //if (compareTime >= DateTime.Now)
                        //{
                        //    Process iisReset = new Process();
                        //    iisReset.StartInfo.Verb = "runas";
                        //    iisReset.StartInfo.FileName = "iisreset.exe";
                        //    iisReset.Start();
                        //    iisReset.WaitForExit();
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                UIException.LogException(ex.Source.ToString(), ex.Message);
            }
            finally
            {
                timer1.Start();
            }
        }

        protected override void OnStop()
        {
        }
        private static void  WriteEventLogEntry(string message,string servicename)
        {
            // Create an instance of EventLog
            System.Diagnostics.EventLog eventLog = new System.Diagnostics.EventLog();

            // Check if the event source exists. If not create it.
            if (!System.Diagnostics.EventLog.SourceExists(servicename))
            {
                System.Diagnostics.EventLog.CreateEventSource(servicename, "Application");
            }

            // Set the source name for writing log entries.
            eventLog.Source = servicename;

            // Create an event ID to add to the event log
            int eventID = 8;

            // Write an entry to the event log.
            eventLog.WriteEntry(message,
                                System.Diagnostics.EventLogEntryType.Information,
                                eventID);

            // Close the Event Log
            eventLog.Close();
        }
    }
}
