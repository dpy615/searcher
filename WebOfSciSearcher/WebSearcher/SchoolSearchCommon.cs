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
                if (!string.IsNullOrEmpty(dt.Rows[i]["isIn"].ToString()) && dt.Rows[i]["isIn"].ToString().ToUpper() != "ERROR") {
                    over++;
                    continue;
                }

                string title_old = dt.Rows[i][0].ToString();
                try {

                    string title = title_old.Replace(' ', '+').Replace("'", "%27").Replace('&', '+');

                    string uri = config.url1 + title + config.url2;

                    WebClient web = new WebClient();
                    string res = web.DownloadString(uri).ToLower().Replace("\r", "").Replace("\n", "");

                    if (config.contentReplace.Count > 0) {
                        foreach (var replace in config.contentReplace) {
                            res = res.Replace(replace[0], replace[1]);
                        }
                    }

                    var collection = Regex.Matches(res, config.titleRegex);
                    if (collection.Count > 0) {
                        double matchValue = 0;
                        string matchTitle = "";
                        string matchWeb = "";

                        #region 找出最匹配的标题
                        foreach (var item in collection) {
                            string str = item.ToString();
                            for (int lIndex = 0; lIndex < config.titleLeft; lIndex++) {
                                str = str.Substring(str.IndexOf('>') + 1);
                            }
                            str = str.Substring(0, str.IndexOf('<') - config.titleRight);
                            if (!string.IsNullOrEmpty(str)) {
                                double dtmp = Utils.MatchValue(title_old, str);
                                if (dtmp > matchValue) {
                                    matchValue = dtmp;
                                    matchTitle = str;
                                    matchWeb = item.ToString();
                                }
                            }
                        }
                        dt.Rows[i]["isIn"] = matchValue.ToString();
                        dt.Rows[i]["matchTitle"] = matchTitle;
                        #endregion

                        #region 判断是否需要转入详细链接，查找更多东西

                        if (matchValue > config.matchGate && !string.IsNullOrEmpty(config.articleRegex)) {
                            //进入详细页面
                            string http = Regex.Match(matchWeb, config.articleRegex).ToString();
                            http = http.Substring(http.IndexOf("\"") + 1);
                            http = http.Substring(0, http.IndexOf("\""));
                            http = config.httpWeb + http;
                            string articleRes = web.DownloadString(http);//.Replace("\r","").Replace("\n","");

                            //是否可以下载
                            if (!string.IsNullOrEmpty(config.downloadRegex)) {
                                var downLoadMatch = Regex.Matches(articleRes, config.downloadRegex);
                                for (int downLoadIndex = 0; downLoadIndex < downLoadMatch.Count && downLoadIndex<6; downLoadIndex++) {
                                    string match = downLoadMatch[downLoadIndex].ToString();
                                    match = match.Substring(match.IndexOf("\"") + 1);
                                    match = config.httpWeb + match.Substring(0, match.IndexOf("\""));
                                    dt.Rows[i]["download" + downLoadIndex] = match;
                                }
                            }
                            //上次修改日期

                        }

                        #endregion


                    } else {
                        dt.Rows[i]["isIn"] = "0";
                    }
                } catch (Exception) {
                    dt.Rows[i]["isIn"] = "Error";
                }

                string strConnection = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source ={0};Extended Properties = Excel 8.0", fileName);
                OleDbConnection oleConnection = new OleDbConnection(strConnection);
                oleConnection.Open();
                string sql = String.Format("update [Sheet1$] set isIn = '{0}',matchTitle = '{1}',download0='{3}',download1='{4}',download2='{5}',download3='{6}',download4='{7}' where 标题 = '{2}'",
                    new string[] { dt.Rows[i]["isIn"].ToString(), 
                        dt.Rows[i]["matchTitle"].ToString(),
                        dt.Rows[i][0].ToString().Replace("'", "''"), 
                        dt.Rows[i]["download0"].ToString(), 
                        dt.Rows[i]["download1"].ToString(), 
                        dt.Rows[i]["download2"].ToString(), 
                        dt.Rows[i]["download3"].ToString(), 
                        dt.Rows[i]["download4"].ToString()
                    });
                OleDbCommand command = new OleDbCommand(sql, oleConnection);
                command.ExecuteNonQuery();
                oleConnection.Close();

                over++;
            }
        }

    }
}
