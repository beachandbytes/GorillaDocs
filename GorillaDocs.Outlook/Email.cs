using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OL = Microsoft.Office.Interop.Outlook;

namespace GorillaDocs.Outlook
{
    public class Email
    {
        public static void Send(string To = null, string Subject = null, string Body = null, string AttachmentPath = null)
        {
            OL.Application app = new OL.Application();
            OL.MailItem mail = app.CreateItem(OL.OlItemType.olMailItem) as OL.MailItem;

            if (!string.IsNullOrEmpty(To))
                mail.To = To;
            if (!string.IsNullOrEmpty(Subject))
                mail.Subject = Subject;
            if (!string.IsNullOrEmpty(Body))
                mail.Body = Body;
            if (!string.IsNullOrEmpty(AttachmentPath))
                mail.Attachments.Add(AttachmentPath);
            mail.Display(false);
        }

        public static void Send(string To = null, string Subject = null, string Body = null, IList<FileInfo> Attachments = null)
        {
            OL.Application app = new OL.Application();
            OL.MailItem mail = app.CreateItem(OL.OlItemType.olMailItem) as OL.MailItem;

            if (!string.IsNullOrEmpty(To))
                mail.To = To;
            if (!string.IsNullOrEmpty(Subject))
                mail.Subject = Subject;
            if (!string.IsNullOrEmpty(Body))
                mail.Body = Body;
            foreach (FileInfo attachment in Attachments)
                mail.Attachments.Add(attachment.FullName);
            mail.Display(false);
        }
    }
}
