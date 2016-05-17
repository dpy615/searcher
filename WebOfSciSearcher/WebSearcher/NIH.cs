using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace WebSearcher {
    public class NIH : Searcher {

        public override void GetRes(object o) {
            ArrayList l = (ArrayList)o;
            int start = (int)l[0];
            int end = (int)l[1];
            for (int i = start; i < end; i++) {
                if (dt.Rows[i]["isIn"].ToString().ToLower() == "yes" || dt.Rows[i]["isIn"].ToString().ToLower() == "no") {
                    over++;
                    continue;
                }
                string title_old = dt.Rows[i][0].ToString();
                try {

                    string title = title_old.Replace(' ', '+');

                    //string uri = "http://eprints.soton.ac.uk/cgi/search/archive/advanced?screen=Search&dataset=archive&documents_merge=ALL&documents=&eprintid=&title_merge=ALL&title=" + title + "&contributors_name_merge=ALL&contributors_name=&abstract_merge=ALL&abstract=&date=&keywords_merge=ALL&keywords=&subjects_merge=ANY&divisions_merge=ANY&department_merge=ALL&department=&refereed=EITHER&publication%2Fseries_name_merge=ALL&publication%2Fseries_name=&documents.date_embargo=&shelves.shelfid=&satisfyall=ALL&order=title%2Fcontributors_name%2F-date&_action_search=Search";
                    string uri = "http://www.ncbi.nlm.nih.gov/pmc/?term=(" + title + ")+AND+%22nih+grants%22%5BFilter%5D";


                    HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(uri);
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    StreamReader reader = new StreamReader(response.GetResponseStream());
                    string res = reader.ReadToEnd();
                    res = DealStr(res);
                    if (res.ToLower().Contains((title_old + "</A>").ToLower())) {
                        dt.Rows[i]["isIn"] = "YES";
                    } else {
                        dt.Rows[i]["isIn"] = "NO";
                    }
                } catch (Exception) {
                    dt.Rows[i]["isIn"] = "Error";
                }

                string strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source =" + fileName + ";Extended Properties = Excel 8.0";
                OleDbConnection oleConnection = new OleDbConnection(strConnection);
                oleConnection.Open();
                string sql = "update [Sheet1$] set isIn = '" + dt.Rows[i]["isIn"] + "' where 标题 = '" + dt.Rows[i][0].ToString().Replace("'", "''") + "'";
                OleDbCommand command = new OleDbCommand(sql, oleConnection);
                command.ExecuteNonQuery();
                oleConnection.Close();

                over++;
                //Console.WriteLine(i);
            }
        }

        private string DealStr(string str_o) {
            string str = str_o;
            str = str.Replace("\r", "").Replace("\n", "").Replace("<B>", "").Replace("</B>", "").Replace("<EM>", "").Replace("</EM>", "");
            str = str.Replace("<b>", "").Replace("</b>", "").Replace("<em>", "").Replace("</em>", "");

            return str;
        }
    }
}
