using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace 百邦所得税汇算底稿工具
{
    public class HtmlTableService
    {
        public DataTable ToDataTalbe(string xmlTableString, Encoding encoding)
        {
            DataTable dt = new DataTable();
            byte[] array = encoding.GetBytes(xmlTableString);
            MemoryStream stream = new MemoryStream(array);
            XElement xml = XElement.Load(stream);
            //检查根节点是否是table
            if (xml.Name.ToString().ToUpper() != "TABLE") return null;
            //获取xml所有tr节点
            var v = xml.Descendants("tr");
            if (v == null || v.Count() == 0) return null;
            int index = 0;
            foreach (var item in v)
            {
                if (index == 0)
                {
//如果是第一行。判断是否有标题 (th)
                    var th = item.Descendants("th");
                    if (th == null || th.Count() == 0)
                    {
//不存在就创建空列
                        var td = item.Descendants("td");
                        dt.Columns.AddRange(CreateDataColumns(td.Count()));
                        //并写入数据
                        WriterDataTable(td.Select(s => s.Value).ToArray(), ref dt);
                    }
                    else
                    {
                        dt.Columns.AddRange(CreateDataColumns(th.Select(s => s.Value).ToArray()));
                    }
                }
                else
                {
                    var td = item.Descendants("td");
                    WriterDataTable(td.Select(s => s.Value).ToArray(), ref dt);
                }
                index++;
            }
            return dt;
        }

        private DataColumn[] CreateDataColumns(string[] columnNames = null)
        {
            DataColumn[] dc = new DataColumn[columnNames.Length];
            for (int i = 0; i < dc.Length; i++)
            {
                dc[i] = new DataColumn(columnNames[i]);
            }
            return dc;
        }

        private DataColumn[] CreateDataColumns(int colCount)
        {
            DataColumn[] dc = new DataColumn[colCount];
            for (int i = 0; i < dc.Length; i++)
            {
                dc[i] = new DataColumn(i.ToString());
            }
            return dc;
        }

        private void WriterDataTable(string[] str, ref DataTable dt)
        {
            DataRow dr = dt.NewRow();
            for (int i = 0; i < (str.Length > dt.Columns.Count ? dt.Columns.Count : str.Length); i++)
            {
                dr[i] = str[i];
            }
            dt.Rows.Add(dr);
        }

        public string[,] ToArray(string xmlTableString, Encoding encoding)
        {
            DataTable dt = ToDataTalbe(xmlTableString, encoding);
            string[,] ss = new string[dt.Rows.Count, dt.Columns.Count];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ss[i, j] = dt.Rows[i][j].ToString();
                }
            }
            return ss;
        }
}

    
}
