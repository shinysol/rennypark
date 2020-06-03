using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Xml;
using System.Xml.Serialization;

namespace MyControls.Helper
{
    public static class XmlHelper
    {
        public async static Task<string> DataTableToXml(this DataTable dt)
        {
            // 2020/06/02 StringWriter 이용해서 DataTable을 Xml 스트링으로 변환
            return await Task.Run(() =>
            {
                using (StringWriter sWriter = new StringWriter())
                using (XmlWriter writer = XmlWriter.Create(sWriter))
                {
                    dt.WriteXml(writer);
                    return sWriter.ToString();
                }
            });
        }
        public async static Task<DataTable> XmlToDataTable(this string xml)
        {
            // 2020/06/02 Xml 스트링을 DataTable로 변환
            return await Task.Run(() =>
            {
                XmlDocument xDoc = new XmlDocument();
                xDoc.LoadXml(xml);
                using (System.Data.DataSet ds = new System.Data.DataSet())
                using (XmlReader xReader = new XmlNodeReader(xDoc))
                {
                    ds.ReadXml(xReader);
                    return ds.Tables.Count > 0 ? ds.Tables[0] : new DataTable();
                }
            });
        }
    }
}
