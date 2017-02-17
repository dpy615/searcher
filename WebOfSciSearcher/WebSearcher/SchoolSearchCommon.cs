using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Data.OleDb;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;
using WebOfSciSearcher;

namespace WebSearcher {
    public class SchoolSearchCommon : Searcher {

        public override void GetRes(object o) {
            ArrayList l = (ArrayList)o;
            int start = (int)l[0];
            int end = (int)l[1];
            float isin;
            for (int i = start; i < end; i++) {
                if (!string.IsNullOrEmpty(dt.Rows[i]["isIn"].ToString()) && float.TryParse(dt.Rows[i]["isIn"].ToString(),out isin)) {
                    over++;
                    continue;
                }

                string title_old = dt.Rows[i][0].ToString();
                try {

                    string title = title_old.Replace(' ', '+').Replace("'", "%27").Replace('&', '+');

                    if (config.urlReplace.Count > 0) {
                        foreach (var replace in config.urlReplace) {
                            title = title.Replace(replace[0], replace[1]);
                        }
                    }


                    string uri = config.url1 + title + config.url2;

                    // HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
                    // string res = new StreamReader(((HttpWebResponse)request.GetResponse()).GetResponseStream()).ReadToEnd();

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
                                double dtmp = Utils.MatchValue(title_old, str.Trim(),false);
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
                                for (int downLoadIndex = 0; downLoadIndex < downLoadMatch.Count && downLoadIndex < 5; downLoadIndex++) {
                                    string match = downLoadMatch[downLoadIndex].ToString();
                                    for (int t = 0; t < config.downloadIndex; t++) {
                                        match = match.Substring(match.IndexOf("\"") + 1);
                                    }
                                    match = config.httpWeb + match.Substring(0, match.IndexOf("\""));
                                    dt.Rows[i]["download" + downLoadIndex] = match;
                                }
                            }

                            if (!string.IsNullOrEmpty(config.detailWeb)) {
                                http = http + config.detailWeb;
                                string details = web.DownloadString(http).Replace("\n", "").Replace("\r", "");
                                Detailes(i, details);
                            }
                        }
                        #endregion


                    } else {
                        dt.Rows[i]["isIn"] = "0";
                    }
                } catch (Exception e) {
                    dt.Rows[i]["isIn"] = e.Message;
                    //MessageBox.Show("读取数据过程中发生错误：" + i + ":" + dt.Rows[i]["标题"].ToString() + "\r\n" + e.ToString());
                }
                try {
                    lock (_locker) {
                        using (var fs = new FileStream(fileName, FileMode.Open)) {
                            HSSFWorkbook workBook = new HSSFWorkbook(fs);
                            ISheet sheet1 = workBook.GetSheet("Sheet1");
                            IRow row = sheet1.GetRow(i + 1);
                            int cellsCount = row.Cells.Count;
                            int columnsCount = sheet1.GetRow(0).Cells.Count;
                            while (cellsCount < columnsCount){
                                row.CreateCell(cellsCount++,CellType.STRING).SetCellValue(" ");
                            }
                            row.Cells[GetColIndex(sheet1, "isIn")].SetCellValue(dt.Rows[i]["isIn"].ToString());
                            row.Cells[GetColIndex(sheet1, "matchTitle")].SetCellValue(dt.Rows[i]["matchTitle"].ToString());
                            row.Cells[GetColIndex(sheet1, "download0")].SetCellValue(dt.Rows[i]["download0"].ToString());
                            row.Cells[GetColIndex(sheet1, "download1")].SetCellValue(dt.Rows[i]["download1"].ToString());
                            row.Cells[GetColIndex(sheet1, "download2")].SetCellValue(dt.Rows[i]["download2"].ToString());
                            row.Cells[GetColIndex(sheet1, "download3")].SetCellValue(dt.Rows[i]["download3"].ToString());
                            row.Cells[GetColIndex(sheet1, "download4")].SetCellValue(dt.Rows[i]["download4"].ToString());
                            row.Cells[GetColIndex(sheet1, "date_accessioned")].SetCellValue(dt.Rows[i]["date_accessioned"].ToString());
                            row.Cells[GetColIndex(sheet1, "date_available")].SetCellValue(dt.Rows[i]["date_available"].ToString());
                            row.Cells[GetColIndex(sheet1, "date_issued")].SetCellValue(dt.Rows[i]["date_issued"].ToString());
                            row.Cells[GetColIndex(sheet1, "language")].SetCellValue(dt.Rows[i]["language"].ToString());
                            row.Cells[GetColIndex(sheet1, "rights")].SetCellValue(dt.Rows[i]["rights"].ToString());
                            row.Cells[GetColIndex(sheet1, "rightsUri")].SetCellValue(dt.Rows[i]["rightsUri"].ToString());
                            row.Cells[GetColIndex(sheet1, "type")].SetCellValue(dt.Rows[i]["type"].ToString());

                            using (var fs1 = new FileStream(fileName, FileMode.Open)) {
                                workBook.Write(fs1);
                            }
                        }
                    }
                } catch (Exception e) {
                    MessageBox.Show("保存数据过程中发生错误：" + i + ":" + dt.Rows[i]["标题"].ToString() + "\r\n" + e.ToString());
                }


                #region 旧的保存方法
                /* 
//                string strConnection = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source ={0};Extended Properties = Excel 8.0", fileName);
//                OleDbConnection oleConnection = new OleDbConnection(strConnection);
//                oleConnection.Open();
//                try {
//                    string sql = String.Format("update [Sheet1$] set isIn = '{0}',matchTitle = '{1}',download0='{3}',download1='{4}',download2='{5}',download3='{6}',download4='{7}' where 标题 = '{2}'",
//                   new string[] { dt.Rows[i]["isIn"].ToString(), 
//                        dt.Rows[i]["matchTitle"].ToString().Replace("'", "''"),
//                        dt.Rows[i][0].ToString().Replace("'", "''"), 
//                        dt.Rows[i]["download0"].ToString(), 
//                        dt.Rows[i]["download1"].ToString(), 
//                        dt.Rows[i]["download2"].ToString(), 
//                        dt.Rows[i]["download3"].ToString(), 
//                        dt.Rows[i]["download4"].ToString()
//                    });
//                    OleDbCommand command = new OleDbCommand(sql, oleConnection);
//                    command.ExecuteNonQuery();
//                } catch (Exception) {
//                    try {
//                        string sql = String.Format("update [Sheet1$] set isIn = '{0}',matchTitle = '{1}',download0='{3}' where 标题 = '{2}'",
//new string[] { dt.Rows[i]["isIn"].ToString(), 
//                        dt.Rows[i]["matchTitle"].ToString().Replace("'", "''"),
//                        dt.Rows[i][0].ToString().Replace("'", "''"), 
//                        "下载链接太长，无法显示"
                        
//                    });
//                        OleDbCommand command = new OleDbCommand(sql, oleConnection);
//                        command.ExecuteNonQuery();
//                    } catch (Exception) {

//                        string sql = String.Format("update [Sheet1$] set isIn = '{0}',matchTitle = '{1}' where 标题 = '{2}'",
//new string[] { dt.Rows[i]["isIn"].ToString(), 
//                        "匹配错误，请手动匹配",
//                        dt.Rows[i][0].ToString().Replace("'", "''")
                        
//                    });
//                        OleDbCommand command = new OleDbCommand(sql, oleConnection);
//                        command.ExecuteNonQuery();
//                    }

//                }


//                oleConnection.Close();
                */
                #endregion
                over++;
            }
        }

        private void Detailes(int i, string details) {
            //上传日期
            if (!string.IsNullOrEmpty(config.dateAssionRegex)) {
                var tmpMatch = Regex.Match(details, config.dateAssionRegex).ToString();
                if (!string.IsNullOrEmpty(tmpMatch)) {
                    for (int li = 0; li < config.dateAssionIndex; li++) {
                        tmpMatch = tmpMatch.Substring(tmpMatch.IndexOf('>') + 1);
                    }
                    tmpMatch = tmpMatch.Substring(0, tmpMatch.IndexOf('<'));
                }
                dt.Rows[i]["date_accessioned"] = tmpMatch;
            }

            //可用日期
            if (!string.IsNullOrEmpty(config.dateAvailableRegex)) {
                var tmpMatch = Regex.Match(details, config.dateAvailableRegex).ToString();
                if (!string.IsNullOrEmpty(tmpMatch)) {
                    for (int li = 0; li < config.dateAvailableIndex; li++) {
                        tmpMatch = tmpMatch.Substring(tmpMatch.IndexOf('>') + 1);
                    }
                    tmpMatch = tmpMatch.Substring(0, tmpMatch.IndexOf('<'));
                }
                dt.Rows[i]["date_available"] = tmpMatch;
            }

            //发表日期
            if (!string.IsNullOrEmpty(config.dateIssuRegex)) {
                var tmpMatch = Regex.Match(details, config.dateIssuRegex).ToString();
                if (!string.IsNullOrEmpty(tmpMatch)) {
                    for (int li = 0; li < config.dateIssuIndex; li++) {
                        tmpMatch = tmpMatch.Substring(tmpMatch.IndexOf('>') + 1);
                    }
                    tmpMatch = tmpMatch.Substring(0, tmpMatch.IndexOf('<'));
                }
                dt.Rows[i]["date_issued"] = tmpMatch;
            }

            //language
            if (!string.IsNullOrEmpty(config.languageRegex)) {
                var tmpMatch = Regex.Match(details, config.languageRegex).ToString();
                if (!string.IsNullOrEmpty(tmpMatch)) {
                    for (int li = 0; li < config.languageIndex; li++) {
                        tmpMatch = tmpMatch.Substring(tmpMatch.IndexOf('>') + 1);
                    }
                    tmpMatch = tmpMatch.Substring(0, tmpMatch.IndexOf('<'));
                }
                dt.Rows[i]["language"] = tmpMatch;
            }

            //rights
            if (!string.IsNullOrEmpty(config.rightsRegex)) {
                var tmpMatch = Regex.Match(details, config.rightsRegex).ToString();
                if (!string.IsNullOrEmpty(tmpMatch)) {
                    for (int li = 0; li < config.rightsIndex; li++) {
                        tmpMatch = tmpMatch.Substring(tmpMatch.IndexOf('>') + 1);
                    }
                    tmpMatch = tmpMatch.Substring(0, tmpMatch.IndexOf('<'));
                }
                dt.Rows[i]["rights"] = tmpMatch;
            }

            //rightsUri
            if (!string.IsNullOrEmpty(config.rightsUriRegex)) {
                var tmpMatch = Regex.Match(details, config.rightsUriRegex).ToString();
                if (!string.IsNullOrEmpty(tmpMatch)) {
                    for (int li = 0; li < config.rightsUriIndex; li++) {
                        tmpMatch = tmpMatch.Substring(tmpMatch.IndexOf('>') + 1);
                    }
                    tmpMatch = tmpMatch.Substring(0, tmpMatch.IndexOf('<'));
                }
                dt.Rows[i]["rightsUri"] = tmpMatch;
            }

            //type
            if (!string.IsNullOrEmpty(config.typeRegex)) {
                var tmpMatch = Regex.Match(details, config.typeRegex).ToString();
                if (!string.IsNullOrEmpty(tmpMatch)) {
                    for (int li = 0; li < config.typeIndex; li++) {
                        tmpMatch = tmpMatch.Substring(tmpMatch.IndexOf('>') + 1);
                    }
                    tmpMatch = tmpMatch.Substring(0, tmpMatch.IndexOf('<'));
                }
                dt.Rows[i]["type"] = tmpMatch;
            }

        }

        public static int GetColIndex(ISheet sheet1, string colName) {
            IRow row = sheet1.GetRow(0);
            for (int i = 0; i < row.Cells.Count; i++) {
                if (row.Cells[i].ToString() == colName) {
                    return i;
                }
            }
            return -1;
        }

    }
}
