using System.IO;
using System;

namespace MyControls
{
    public class Filecontrol : IDisposable
    {
        protected string filepath;
        bool disposed = false;
        System.Runtime.InteropServices.SafeHandle handle = new Microsoft.Win32.SafeHandles.SafeFileHandle(IntPtr.Zero, true);  // Dispose
        public string Filepath
        {
            get { return filepath; }
            set { filepath = value; }
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        public Filecontrol(string path)
        {
            Filepath = path;
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
        ~Filecontrol()
        {
            Dispose(false);
        }
        //============IDisposable============

        public void StreamWrite(string content)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(filepath));
            FileStream fs = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Write);
            using (StreamWriter writer = new StreamWriter(fs))
            {
                writer.Write(Crypto.Encrypt(content));
            }
        }
        public string StreamRead()
        {
            FileStream fs = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Read);
            using (StreamReader reader = new StreamReader(fs))
            {
                return Crypto.Decrypt(reader.ReadToEnd());
            }
        }
    }
}