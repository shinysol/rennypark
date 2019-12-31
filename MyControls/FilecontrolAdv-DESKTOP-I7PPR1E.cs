using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls
{
    public class FilecontrolAdv : Filecontrol
    {
        string oripath;
        public FilecontrolAdv(string path) : base(path)
        {
            oripath = path;
        }
        public void StreamWriteLocalAppdata(string content)
        {
            filepath = Path.Combine(Environment.GetEnvironmentVariable("LocalAppData"), oripath);
            StreamWrite(content);
        }
        public void StreamWriteDesktop(string content)
        {
            filepath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), oripath);
            StreamWrite(content);
        }
    }
}
