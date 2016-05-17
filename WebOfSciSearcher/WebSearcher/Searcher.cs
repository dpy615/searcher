using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WebSearcher {
    public abstract class Searcher {
        public DataTable dt = new DataTable();
        public List<Thread> threads = new List<Thread>();
        public double over = 0;
        public double all = 0;
        public bool isRun = true;
        public string fileName = "";

        public int yes = 0;
        public int no = 0;
        public int error = 0;


        public void GetData(int thCount, string fileName) {
            this.fileName = fileName;
            if (thCount < 1) return;
            string strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source =" + fileName + ";Extended Properties = Excel 8.0";
            OleDbConnection oleConnection = new OleDbConnection(strConnection);
            OleDbDataAdapter adapter = new OleDbDataAdapter("select * from [Sheet1$]", oleConnection);
            adapter.Fill(dt);
            oleConnection.Close();
            if (dt.Rows.Count < 1) {
                return;
            }
            all = dt.Rows.Count;

            int count = dt.Rows.Count / thCount;
            for (int i = 0; i < thCount && isRun; i++) {
                ArrayList l = new ArrayList();
                l.Add(i * count);
                if (i == thCount - 1) {
                    l.Add(dt.Rows.Count);
                } else {
                    l.Add((i + 1) * count);
                }
                Thread t = new Thread(new ParameterizedThreadStart(GetRes));
                t.Start(l);
                threads.Add(t);
            }
        }

        public abstract void GetRes(object o);
    }
}
