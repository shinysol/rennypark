using System;
using System.Data;
using System.Drawing;
using System.Net;
using System.Configuration;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using System.Collections.Generic;

namespace MyControls
{
    public class Smtpcontrol : IDisposable
    {
        bool disposed = false;
        SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);  // Dispose
                                                                    // Public implementation of Dispose pattern callable by consumers.
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
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
        NetworkCredential basicCredential;
        public Smtpcontrol(string ID, string PW)
        {
            basicCredential = new NetworkCredential(ID, PW);
        }
        ~Smtpcontrol()
        {
            Dispose(false);
        }
        public bool SendMessage(string from, string to, bool htmlorNot, string htmlpathorBody)
        {
            string mailbody;
            if (htmlorNot)
            {
                string filename = @"D:\Event.html";
                mailbody = System.IO.File.ReadAllText(filename);
            }
            else
            {
                mailbody = htmlpathorBody;
            }
            MailMessage message = new MailMessage(from, to)
            {
                Subject = "Auto Response Email",
                Body = mailbody,
                BodyEncoding = Encoding.UTF8,
                IsBodyHtml = htmlorNot,
            };
            Attachment attachment;
            attachment = new Attachment("your attachment file");
            message.Attachments.Add(attachment);
            
            SmtpClient client = new SmtpClient("webmail.aonecustoms.com", 587)
            {
                UseDefaultCredentials = true,
                Credentials = basicCredential,
            };
            try
            {
                client.Send(message);
                return true;
            }
            catch
            {
                return false;
            }
        }
        
    }
}
