using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleCalenderAssist
{
    class XML
    {
        /// <summary>
        /// 開く
        /// </summary>
        public DataSet Open(string fname)
        {
            using (DataSet ds = new DataSet())
            {
                ds.ReadXml(fname);
                return ds;
            }
        }

        /// <summary>
        /// 保存
        /// </summary>
        public void Save(string fname, DataSet ds)
        {
            using(StreamWriter sWrite = new StreamWriter(fname, false, System.Text.Encoding.Default))
            {
                ds.WriteXml(sWrite, XmlWriteMode.WriteSchema);
            }
        }
    }
}
