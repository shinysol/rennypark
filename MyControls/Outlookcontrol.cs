using System;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MyControls
{
    public class Outlookcontrol : IDisposable
    {
        bool disposed = false;
        SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);
        public void Send(string Recipients, string Subject, string Body)
        {
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mail = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.Subject = Subject;
            mail.Body = Body;
            Outlook.AddressEntry currentUser = app.Session.CurrentUser.AddressEntry;
            //Outlook.ExchangeUser manager = currentUser.GetExchangeUser().GetExchangeUserManager();
            // Add recipient using display name, alias, or smtp address
            mail.Recipients.Add(Recipients);
            mail.Recipients.ResolveAll();
            mail.Send();
        }
        public void SendHTML(string Recipients, string Subject, string Body)
        {
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mail = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.Subject = Subject;
            mail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            mail.HTMLBody = Body;
            Outlook.AddressEntry currentUser = app.Session.CurrentUser.AddressEntry;
            //Outlook.ExchangeUser manager = currentUser.GetExchangeUser().GetExchangeUserManager();
            // Add recipient using display name, alias, or smtp address
            mail.Recipients.Add(Recipients);
            mail.Recipients.ResolveAll();
            mail.Send();
        }
        public void Send(string Recipients, string Subject, string Body, string AttachmentPath)
        {
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mail = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.Subject = Subject;
            mail.Body = Body;
            Outlook.AddressEntry currentUser = app.Session.CurrentUser.AddressEntry;
            //Outlook.ExchangeUser manager = currentUser.GetExchangeUser().GetExchangeUserManager();
            // Add recipient using display name, alias, or smtp address
            mail.Recipients.Add(Recipients);
            mail.Recipients.ResolveAll();
            mail.Attachments.Add(AttachmentPath, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
            mail.Send();
        }
        public void Send(string[] Recipients, string Subject, string Body)
        {
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mail = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.Subject = Subject;
            mail.Body = Body;
            Outlook.AddressEntry currentUser = app.Session.CurrentUser.AddressEntry;
            //Outlook.ExchangeUser manager = currentUser.GetExchangeUser().GetExchangeUserManager();
            // Add recipient using display name, alias, or smtp address
            foreach (string str in Recipients)
            {
                if (str != string.Empty)
                {
                    mail.Recipients.Add(str);
                }
            }
            mail.Recipients.ResolveAll();
            mail.Send();
        }
        public void Send(string[] Recipients, string Subject, string Body, string AttachmentPath)
        {
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mail = app.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.Subject = Subject;
            mail.Body = Body;
            Outlook.AddressEntry currentUser = app.Session.CurrentUser.AddressEntry;
            //Outlook.ExchangeUser manager = currentUser.GetExchangeUser().GetExchangeUserManager();
            // Add recipient using display name, alias, or smtp address
            foreach (string str in Recipients)
            {
                if (str != string.Empty)
                {
                    mail.Recipients.Add(str);
                }
            }
            mail.Recipients.ResolveAll();
            mail.Attachments.Add(AttachmentPath, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
            mail.Send();
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        public virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                handle.Dispose();
                // Free any other managed objects here.
                //
            }

            // Free any unmanaged objects here.
            //
            disposed = true;
        }
        ~Outlookcontrol()
        {
            Dispose(false);
        }
    }
}