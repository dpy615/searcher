using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WebOfSciSearcher;

namespace WebSearcher {
    public abstract class Searcher {
        public DataTable dt = new DataTable();
        public List<Thread> threads = new List<Thread>();
        public double over = 0;
        public double all = 0;
        public bool isRun = true;
        public string fileName = "";
        public Config config;

        public int yes = 0;
        public int no = 0;
        public int error = 0;

        public static object _locker = new object();


        public void GetData(int thCount, string fileName) {
            this.fileName = fileName;

            CheckCol("isIn");
            CheckCol("matchTitle");
            CheckCol("date_accessioned");
            CheckCol("date_available");
            CheckCol("date_issued");
            CheckCol("language");
            CheckCol("rights");
            CheckCol("rightsUri");
            CheckCol("type");
            CheckCol("downloadType");
            CheckCol("download0");
            CheckCol("download1");
            CheckCol("download2");
            CheckCol("download3");
            CheckCol("download4");


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

        public void CheckCol(string colName) {
            lock (_locker) {
                string strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source =" + fileName + ";Extended Properties = Excel 8.0";
                OleDbConnection oleConnection = new OleDbConnection(strConnection);
                oleConnection.Open();
                try {
                    OleDbDataAdapter adapter = new OleDbDataAdapter("select " + colName + " from [Sheet1$]", oleConnection);
                    adapter.Fill(new DataTable());
                } catch (Exception) {
                    oleConnection.Close();
                    var fs = new FileStream(fileName, FileMode.Open);
                    HSSFWorkbook workBook = new HSSFWorkbook(fs);
                    ISheet sheet1 = workBook.GetSheet("Sheet1");
                    IRow row = sheet1.GetRow(0);
                    row.CreateCell(row.Cells.Count).SetCellValue(colName);

                    var fs1 = new FileStream(fileName, FileMode.Open);
                    workBook.Write(fs1);
                    fs1.Close();
                } finally {
                    try {
                        oleConnection.Close();
                    } catch (Exception) {
                    }
                }
            }

        }
    }
}
