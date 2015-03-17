using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OL = Microsoft.Office.Interop.Outlook;

namespace GorillaDocs.Outlook
{
    public class Support
    {
        public static void SendEmail(string To = null, string Subject = null, string Body = null, string LogFileFullName = null)
        {
            Subject = string.IsNullOrEmpty(Subject) ? "Template Support: " : Subject;
            Body = string.IsNullOrEmpty(Body) ? "Please enter a relevant subject and a detailed description of your support item below (include screenshots and outline the steps to reproduce the issue where possible):" : Body;
            Email.Send(To, Subject, Body, LogFileFullName);
        }
        public static void SendEmail(FileInfo logfile, string To = null) { SendEmail(To, null, null, logfile.FullName); }
    }
}
