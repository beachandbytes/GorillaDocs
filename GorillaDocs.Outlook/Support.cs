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
            OL.Application app = new OL.Application();
            OL.MailItem mail = app.CreateItem(OL.OlItemType.olMailItem) as OL.MailItem;

            if (!string.IsNullOrEmpty(To))
                mail.To = To;
            mail.Subject = string.IsNullOrEmpty(Subject) ? "Template Support: " : Subject;
            mail.Body = string.IsNullOrEmpty(Body) ? "Please enter a relevant subject and a detailed description of your support item below (include screenshots and outline the steps to reproduce the issue where possible):" : Body;
            if (!string.IsNullOrEmpty(LogFileFullName))
                mail.Attachments.Add(LogFileFullName);
            mail.Display(false);
        }
        public static void SendEmail(FileInfo logfile) { SendEmail(null, null, null, logfile.FullName); }
    }
}
