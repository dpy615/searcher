using System;
using System.Collections;
using System.Data.OleDb;
using System.Net;
using System.Text.RegularExpressions;
using WebOfSciSearcher;

namespace WebSearcher {
    public class SchoolSearchCommon : Searcher {

        public override void GetRes(object o) {
            ArrayList l = (ArrayList)o;
            int start = (int)l[0];
            int end = (int)l[1];
            for (int i = start; i < end; i++) {
                if (!string.IsNullOrEmpty(dt.Rows[i]["isIn"].ToString())) {
                    over++;
                    continue;
                }

                string title_old = dt.Rows[i][0].ToString();
                try {

                    string title = title_old.Replace(' ', '+').Replace("'", "%27").Replace('&','+');

                    string uri = config.url1+title+config.url2;

                    WebClient web = new WebClient();
                    string res = web.DownloadString(uri).ToLower();

                    var collection = Regex.Matches(res, config.titleRegex);
                    if (collection.Count > 0) {
                        double matchValue = 0;
                        string matchTitle = "";
                        string matchWeb = "";
                        foreach (var item in collection) {
                            string str = item.ToString();
                            for (int lIndex = 0; lIndex < config.titleLeft; lIndex++) {
                                str = str.Substring(str.IndexOf('>')+1);
                            }
                            str = str.Substring(0, str.IndexOf('<')-config.titleRight);
                            if (!string.IsNullOrEmpty(str)) {
                                double dtmp =Utils.MatchValue(title_old, str);
                                if (dtmp > matchValue) {
                                    matchValue = dtmp;
                                    matchTitle = str;
                                }
                            }
                        }
                        dt.Rows[i]["isIn"] = matchValue.ToString();
                        dt.Rows[i]["matchTitle"] = matchTitle;
                    } else {
                        dt.Rows[i]["isIn"] = "0";
                    }
                } catch (Exception) {
                    dt.Rows[i]["isIn"] = "Error";
                }

                string strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source =" + fileName + ";Extended Properties = Excel 8.0";
                OleDbConnection oleConnection = new OleDbConnection(strConnection);
                oleConnection.Open();
                string sql = "update [Sheet1$] set isIn = '" + dt.Rows[i]["isIn"] + "',matchTitle = '"+dt.Rows[i]["matchTitle"]+"' where 标题 = '" + dt.Rows[i][0].ToString().Replace("'", "''") + "'";
                OleDbCommand command = new OleDbCommand(sql, oleConnection);
                command.ExecuteNonQuery();
                oleConnection.Close();

                over++;
            }
        }
        
    }
}
