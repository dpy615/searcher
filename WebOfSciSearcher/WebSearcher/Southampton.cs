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
    public class Southampton :Searcher {

        public override void GetRes(object o) {
            ArrayList l = (ArrayList)o;
            int start = (int)l[0];
            int end = (int)l[1];
            for (int i = start; i < end; i++) {
                if (dt.Rows[i]["isIn"].ToString().ToLower() == "yes") {
                    over++;
                    yes++;
                    continue;
                }
                if (dt.Rows[i]["isIn"].ToString().ToLower() == "no") {
                    over++;
                    no++;
                    continue;
                }
                string title_old = dt.Rows[i][0].ToString();
                try {

                    string title = title_old.Replace(' ', '+').Replace("'", "%27").Replace('&','+');

                    string uri = "http://eprints.soton.ac.uk/cgi/search/archive/advanced?screen=Search&dataset=archive&documents_merge=ALL&documents=&eprintid=&title_merge=ALL&title="+title+"&contributors_name_merge=ALL&contributors_name=&abstract_merge=ALL&abstract=&date=&keywords_merge=ALL&keywords=&subjects_merge=ANY&divisions_merge=ANY&department_merge=ALL&department=&refereed=EITHER&publication%2Fseries_name_merge=ALL&publication%2Fseries_name=&documents.date_embargo=&shelves.shelfid=&satisfyall=ALL&order=title%2Fcontributors_name%2F-date&_action_search=Search";

                    HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(uri);
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    StreamReader reader = new StreamReader(response.GetResponseStream());
                    string res = reader.ReadToEnd();
                    if (res.Contains("Search has no matches")) {
                        dt.Rows[i]["isIn"] = "NO";
                    } else {

                        //string regex = "<a href=\"http://eprints.soton.ac.uk/[0-9]*/\">" + strToReg(title_old);

                        //string t = Regex.Match(res, regex).ToString();
                        //string[] webs = t.Split('"');
                        //string newUri = webs[1];
                        //WebClient wc = new WebClient();
                        //StreamReader sr = new StreamReader( wc.OpenRead(newUri));
                        dt.Rows[i]["isIn"] = "YES";
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
            }
        }
        public static string strToReg(string str) {
            string sreturn = "";
            sreturn = str.Replace("[", "\\[");
            sreturn = sreturn.Replace("]", "\\]");
            sreturn = sreturn.Replace("{", "\\{");
            sreturn = sreturn.Replace("}", "\\}");
            sreturn = sreturn.Replace("(", "\\(");
            sreturn = sreturn.Replace(")", "\\)");
            sreturn = sreturn.Replace("\"", "\\\"");
            sreturn = sreturn.Replace("'", "\\'");
            sreturn = sreturn.Replace("?", "\\?");
            sreturn = sreturn.Replace(".", "\\.");
            sreturn = sreturn.Replace(",", "\\,");
            sreturn = sreturn.Replace("|", "\\|");
            sreturn = sreturn.Replace("*", "\\*");
            sreturn = sreturn.Replace("!", "\\!");
            sreturn = sreturn.Replace("~", "\\~");
            sreturn = sreturn.Replace("#", "\\#");
            sreturn = sreturn.Replace("^", "\\^");
            sreturn = sreturn.Replace("&", "\\&");
            sreturn = sreturn.Replace("+", "\\+");
            sreturn = sreturn.Replace("-", "\\-");
            return sreturn;
        }
    }
}
