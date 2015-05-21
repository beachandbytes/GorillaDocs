using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace GorillaDocs
{
    public class ProblemStepsRecorder
    {
        public static void Start(FileInfo file)
        {
            if (file.Exists)
                file.Delete();

            if (Process.GetProcessesByName("psr").Any())
                Process.GetProcessesByName("psr").First().Kill();

            using (var process = new Process())
            {

                ProcessStartInfo startInfo = new ProcessStartInfo()
                {
                    WindowStyle = ProcessWindowStyle.Hidden,
                    FileName = "psr.exe",
                    Arguments = string.Format(@"/start /output ""{0}"" /gui 0 /sc 1 /arcxml 1 /maxlogsize 10", file.FullName)
                };
                process.StartInfo = startInfo;
                process.Start();
            }
        }

        public static void Stop()
        {
            using (var process = new Process())
            {
                ProcessStartInfo startInfo = new ProcessStartInfo()
                {
                    WindowStyle = ProcessWindowStyle.Hidden,
                    FileName = "psr.exe",
                    Arguments = "/stop"
                };
                process.StartInfo = startInfo;
                process.Start();
            }
        }
    }
}
