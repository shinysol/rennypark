using System;
using System.Xml;
using System.Text;
using System.Data;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;

namespace MyControls
{
    public class Xmlcontrol : IDisposable
    {
        bool disposed = false;
        SafeHandle handle = new SafeFileHandle(IntPtr.Zero, true);

        //public DataSet API001(string key, string BlNo, string year)
        //{
        //    // 화물진행정보 조회!
        //    DataSet ds = new DataSet();
        //    StringBuilder st = new StringBuilder();
        //    st.Append("https://unipass.customs.go.kr:38010/ext/rest/cargCsclPrgsInfoQry/retrieveCargCsclPrgsInfo?crkyCn=");
        //    st.Append(key);
        //    st.Append("&hblNo=");
        //    st.Append(BlNo);
        //    st.Append("&blYy");
        //    st.Append(year);
        //    XmlDocument xml = OpenXml(st.ToString());
        //    XmlReader xr = new XmlNodeReader(xml);
        //    ds.ReadXml(xr);
        //    return ds;
        //}
        //public DataSet API005(string key, string abc)
        //{
        //    DataSet ds = new DataSet();
        //    return ds;
        //}


        protected XmlDocument OpenXml(string path)
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(path);
            return xml;
        }
        protected void CloseXml(XmlDocument xml)
        {
            //필요한가?
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
        ~Xmlcontrol()
        {
            Dispose(false);
        }
    }
}