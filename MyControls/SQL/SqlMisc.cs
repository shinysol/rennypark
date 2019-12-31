using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyControls
{
    static class SqlMisc
    {
        //Sql Class를 만든다.
        public static string DbString(string str)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("'");
            sb.Append(str);
            sb.Append("'");
            return sb.ToString();
        }
        public static string DbEqualString(string str1, string str2)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(str1);
            sb.Append(" = ");
            sb.Append(str2);
            return sb.ToString();
        }
        public static string DbInString(string str1, string str2, bool last)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(str1);
            sb.Append(" in ");
            sb.Append(str2);
            if (!last)
            {
                sb.Append(",");
            }
            return sb.ToString();
        }
        public static string DbLsString(string str1, string str2, bool last)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(str1);
            sb.Append(" < ");
            sb.Append(str2);
            if (!last)
            {
                sb.Append(",");
            }
            return sb.ToString();
        }
        public static string DbLtString(string str1, string str2, bool last)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(str1);
            sb.Append(" <= ");
            sb.Append(str2);
            if (!last)
            {
                sb.Append(",");
            }
            return sb.ToString();
        }
        public static string DbBetString(string str1, string bet1, string bet2, bool last)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(str1);
            sb.Append(" between ");
            sb.Append(bet1);
            sb.Append(" and ");
            sb.Append(bet2);
            if (!last)
            {
                sb.Append(",");
            }
            return sb.ToString();
        }
        public static string DbInsertString(string str, bool last)
        {
            StringBuilder st = new StringBuilder();
            st.Append("'");
            if (!str.Equals(string.Empty))
            {
                st.Append(str);
            }
            st.Append("'");
            if (!last)
            {
                st.Append(",");
            }
            return st.ToString();
        }
        public static string DbInsertInt(int str, bool last)
        {
            StringBuilder st = new StringBuilder();
            st.Append(str);
            if (!last)
            {
                st.Append(",");
            }
            return st.ToString();
        }
        public static string DbInsertInt(string str, bool last)
        {
            StringBuilder st = new StringBuilder();
            if (!str.Equals(string.Empty))
            {
                st.Append(str);
            }
            else
            {
                st.Append("null");
            }
            if (!last)
            {
                st.Append(",");
            }
            return st.ToString();
        }
        public static string DbUpdateInt(string col, string value)
        {
            StringBuilder sb = new StringBuilder();
            if (TypeMisc.IsTextAllowed(value))
            {
                sb.Append(" ");
                sb.Append(col);
                sb.Append(" = ");
                sb.Append(value);
                sb.Append(",");
                return sb.ToString();
            }
            return null;
        }
        public static string DbUpdateInt(string col, string value, bool last)
        {
            if (value.Equals(string.Empty))
            {
                return string.Empty;
            }
            StringBuilder st = new StringBuilder();
            st.Append(col);
            st.Append(" = ");
            st.Append(value);
            if (!last)
            {
                st.Append(",");
            }
            return st.ToString();
        }
        public static string DbUpdateCur(string col, string value)
        {
            StringBuilder sb = new StringBuilder();
            string conv = value.Replace(",", "");
            if (TypeMisc.IsTextAllowed(conv))
            {
                sb.Append(" ");
                sb.Append(col);
                sb.Append(" = ");
                sb.Append(conv);
                sb.Append(",");
                return sb.ToString();
            }
            return null;
        }
        public static string DbUpdateString(string col, string value)
        {
            if (value != string.Empty)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(" ");
                sb.Append(col);
                sb.Append(" = '");
                sb.Append(value);
                sb.Append("',");
                return sb.ToString();
            }
            return null;
        }
        public static string DbUpdateString(string col, string value, bool last)
        {
            StringBuilder st = new StringBuilder();
            st.Append(col);
            st.Append(" = ");
            if (value.Equals(string.Empty))
            {
                st.Append("null");
            }
            else
            {
                st.Append("'");
                st.Append(value);
                st.Append("'");
            }
            if (!last)
            {
                st.Append(",");
            }
            return st.ToString();
        }
        public static string DbUpdateString(string tableName, Dictionary<string, bool> columnValues, List<string> where)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Update ");
            sb.Append(tableName);
            sb.Append(" set ");
            foreach (KeyValuePair<string, bool> cv in columnValues)
            {
                sb.Append(cv.Key);
                sb.Append(" = ");
                sb.Append(cv.Value);
                sb.Append(",");
            }
            sb.Remove(sb.Length - 1, 1);
            if (!ReferenceEquals(where, null))
            {
                sb.Append(" where ");
                foreach (string str in where)
                {
                    sb.Append(str);
                    sb.Append(" and ");
                }
                sb.Remove(sb.Length - 5, 5);
            }
            return sb.ToString();
        }
        public static string DbSelectString(string tableName, List<string> columns, List<string> where, Dictionary<string, bool> orderBy)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Select ");
            foreach (string str in columns)
            {
                sb.Append(str);
                sb.Append(",");
            }
            
            sb.Remove(sb.Length - 1, 1);
            sb.Append(" from ");
            sb.Append(tableName);
            if (!ReferenceEquals(where, null))
            {
                sb.Append(" where ");
                foreach (string str in where)
                {
                    sb.Append(str);
                    sb.Append(" and ");
                }
                sb.Remove(sb.Length - 5, 5);
            }
            if (!ReferenceEquals(orderBy, null))
            {
                sb.Append(" Order by ");
                foreach (KeyValuePair<string, bool> ord in orderBy)
                {
                    sb.Append(ord.Key);
                    if (ord.Value)
                    {
                        sb.Append(" ASC");
                    }
                    else
                    {
                        sb.Append(" DESC");
                    }
                    sb.Append(",");
                }
                sb.Remove(sb.Length - 1, 1);
            }
            return sb.ToString();
        }
    }
}
