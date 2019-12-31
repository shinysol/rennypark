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
    class Pop3control : IDisposable
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
        public Pop3control(string ID, string PW)
        {
            basicCredential = new NetworkCredential(ID, PW);
            
        }
        ~Pop3control()
        {
            Dispose(false);
        }
    }
}
