using System;
using McMaster.Extensions.CommandLineUtils;
using Microsoft.Win32.TaskScheduler;

namespace FixTime
{
    class Program
    {
        static void Main(string[] args) => CommandLineApplication.Execute<Program>(args);

        [Option(Template = "-x", Description = "This flag turns off scheduling a monthly task to reregister this process", ShortName = "X")]
        public bool DontScheduleTask { get; }

        private void OnExecute()
        {
            if (System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows) && IsAdministrator())
            {
                RegisterNTPServer();

                if (!DontScheduleTask)
                {
                    ScheduleTask();
                }
            }
            else
            {
                Console.WriteLine("This application is only for Windows operating Systems and has to be run as an Administrator");
                Console.ReadLine();
            }
        }

        private static bool IsAdministrator()
        {
            return (new System.Security.Principal.WindowsPrincipal(System.Security.Principal.WindowsIdentity.GetCurrent())).IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
        }  

        private void RegisterNTPServer()
        {
            var NetStopProcess = System.Diagnostics.Process.Start("net.exe", "stop w32time");
            NetStopProcess.WaitForExit();

            var UnregisterProcess = System.Diagnostics.Process.Start("w32tm.exe", "/unregister");
            UnregisterProcess.WaitForExit();

            var RegisterProcess = System.Diagnostics.Process.Start("w32tm.exe", "/register");
            RegisterProcess.WaitForExit();

            var NetStartProcess = System.Diagnostics.Process.Start("net.exe", "start w32time");
            NetStartProcess.WaitForExit();
        }

        private void ScheduleTask()
        {
            if (CheckTemp())
            {
                Console.WriteLine("Folder Structure has been made at C:\\Temp");

                if (WriteScriptToFile())
                {
                    Console.WriteLine("Scripts have been written to file");

                    if (RunScripts())
                    {
                        Console.WriteLine("Monthly Schedule has been installed");
                        Console.ReadLine();
                    }
                }
            }
        }

        private bool CheckTemp()
        {
            System.IO.DirectoryInfo di = System.IO.Directory.CreateDirectory(@"C:\Temp\FixTime");

            if (di.Exists)
            {
                System.IO.DirectoryInfo diTemp = new System.IO.DirectoryInfo(@"C:\Temp");

                //See if directory has hidden flag, if not, make hidden
                if ((diTemp.Attributes & System.IO.FileAttributes.Hidden) != System.IO.FileAttributes.Hidden)
                {   
                    //Add Hidden flag    
                    diTemp.Attributes |= System.IO.FileAttributes.Hidden;    
                }

                return true;
            }
            else
            {
                return false;
            }
        }

        private string content = @"net stop w32time
        w32tm /unregister
        w32tm /register
        net start w32time";

        private string filename = @"C:\Temp\FixTime\Fixtime.cmd";

        private bool WriteScriptToFile()
        {
            if (!System.IO.File.Exists(filename))
            {
                try
                {
                	using(System.IO.StreamWriter sw = System.IO.File.CreateText(filename))
                	{
                    		sw.WriteLine(content);
                	}
		        }
                catch(Exception e)
                {
                    Console.WriteLine($"Couldn't write to path: { filename }{ Environment.NewLine }{ e }");
                    return false;
                }

                return true;
            }

            return false;
        }

        private bool RunScripts()
        {
            try
            {
                using (TaskService ts = new TaskService())
                {
                    TaskDefinition td = ts.NewTask();
                    td.RegistrationInfo.Description = "resets ntp server once a month";

                    td.Triggers.Add(new MonthlyTrigger());

                    td.Actions.Add(new ExecAction(@"C:\Temp\FixTime\FixTime.cmd"));

                    td.Principal.RunLevel = TaskRunLevel.Highest;

                    ts.RootFolder.RegisterTaskDefinition(@"Fix Time", td, TaskCreation.CreateOrUpdate, "SYSTEM", null, TaskLogonType.ServiceAccount);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Couldn't create Scheduled Task{ Environment.NewLine }{ e }");
                Console.ReadLine();
                return false;
            }
            
            return true;
        }
    }
}
