using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace MyControls
{
    // 둘 이상의 인스턴스가 동시에 read/write를 할 경우에 어떻게 예외처리를 해야 하는가?
    public enum IOcontrolPathType
    {
        FullPath = 1,
        AppData = 2,
        DeskTop = 3
    }
    public enum IOcontrolRandomType
    {
        InAppData = 1,
    }
    public class IOcontrol : IDisposable
    {
        private string filePathStr;
        public string FilePath { get => filePathStr; set => filePathStr = value;}
        bool disposed = false;
        System.Runtime.InteropServices.SafeHandle handle = new Microsoft.Win32.SafeHandles.SafeFileHandle(IntPtr.Zero, true);  // Dispose
        public bool CheckFileExists()
        {
            return File.Exists(filePathStr);
        }
        public void StreamWrite(string content)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(filePathStr));
                FileStream fs = new FileStream(filePathStr, FileMode.OpenOrCreate, FileAccess.Write);
                using (StreamWriter writer = new StreamWriter(fs))
                {
                    writer.Write(content);
                }
            }
            catch(Exception ex)
            {
                Debug.Print(ex.ToString());
            }
        }
        public async Task StreamWriteAsync(string content)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(filePathStr));
                FileStream fs = new FileStream(filePathStr, FileMode.Create, FileAccess.Write);
                using (StreamWriter writer = new StreamWriter(fs))
                {
                    await writer.WriteAsync(content);
                }
            }
            catch(Exception ex)
            {
                Debug.Print(ex.ToString());
            }
        }
        public string StreamRead()
        {
            try
            {
                FileStream fs = new FileStream(filePathStr, FileMode.OpenOrCreate, FileAccess.Read);
                using (StreamReader reader = new StreamReader(fs))
                {
                    return reader.ReadToEnd();
                }
            }
            catch(Exception ex)
            {
                Debug.Print(ex.ToString());
                return string.Empty;
            }
        }
        public async Task<string> StreamReadAsync()
        {
            try
            {
                FileStream fs = new FileStream(filePathStr, FileMode.OpenOrCreate, FileAccess.Read);
                using (StreamReader reader = new StreamReader(fs))
                {
                    return await reader.ReadToEndAsync();
                }
            }
            catch(Exception ex)
            {
                Debug.Print(ex.ToString());
                return string.Empty;
            }
        }
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
        ~IOcontrol()
        {
            Dispose(false);
        }
        // 생성자
        public IOcontrol(IOcontrolPathType PathType) : this(string.Empty, PathType)
        {
            // 상속받을 클래스용 생성자
            OnIOcontrol(PathType);
        }
        protected virtual void OnIOcontrol(IOcontrolPathType PathType)
        {

        }
        public IOcontrol(string filePath, IOcontrolPathType PathType)
        {
            switch (PathType)
            {
                case IOcontrolPathType.FullPath:
                    filePathStr = filePath;
                    break;
                case IOcontrolPathType.AppData:
                    filePathStr = Path.Combine(Environment.GetEnvironmentVariable("LocalAppData"), filePath);
                    break;
                case IOcontrolPathType.DeskTop:
                    filePathStr = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), filePath);
                    break;
            }
        }
        public IOcontrol(IOcontrolRandomType RandomType)
        {
            switch (RandomType)
            {
                case IOcontrolRandomType.InAppData:
                    filePathStr = Path.Combine(Environment.GetEnvironmentVariable("LocalAppData"), Path.GetRandomFileName());
                    break;
            }
        }
    }
}