using System;
using System.IO;

namespace MyControls
{
    class FilecontrolAdvCfg
    {
        private string oriPath;
        string filePath;
        public FilecontrolAdvCfg(string configFileName) => oriPath = configFileName;
        CryptoAdv CrtA = new CryptoAdv(@"A@neCust0ms", @"S@ltK2y", @"1A2b@3C4d#5e6F!0");
        public void StreamWrite(string content)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(filePath));
            FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
            using (StreamWriter writer = new StreamWriter(fs))
            {
                writer.Write(CrtA.Encrypt(content));
            }
        }
        public string StreamRead()
        {
            FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Read);
            using (StreamReader reader = new StreamReader(fs))
            {
                return CrtA.Decrypt(reader.ReadToEnd());
            }
        }
        public void StreamWriteLocalAppdata(string content)
        {
            filePath = Path.Combine(Environment.GetEnvironmentVariable("LocalAppData"), oriPath);
            StreamWrite(content);
        }
        public void StreamWriteDesktop(string content)
        {
            filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), oriPath);
            StreamWrite(content);
        }
    }
}