using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls
{
    public class Structs
    {
        public struct SqlVar
        {
            public string varNm;
            public string udtNm;
            public object varVal;
            public SqlDbType varType;
            public int varSize;
            public System.Windows.Controls.Control varControl;
        }
        public struct LbxSelectedItem
        {
            public int SN;
            public string SelectedString;
        }
    }
}