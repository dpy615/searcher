using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.OleDb;
using System.Threading;

namespace MySearcher {
    public class HtmAnalyse {

        public DataTable dt = new DataTable();

        public int step;
        public int count;
        string folderName;
        public HtmAnalyse(string folderName, int count) {
            this.folderName = folderName;
            this.count = count;
            this.step = 0;
        }

        public static string reg_all = @"[#@￥\""\[\]\{\}\&\;\!\+\=\*\?\'\,\(\)\./:_\- a-zA-Z0-9]";

        public static string reg_PageNo = "<input class=\"goToPageNumber-input\" type=\"text\" name=\"page\" value=\"\\d\" size=\"5\" style=\"width:40px;\" />";
        public static string reg_PageCount = "<span id=\"pageCount.bottom\">[\\,\\d]*</span>";
        public static string reg_title = @"<value lang_id=\""\"">" + reg_all + @"*(<span class=\""hitHilite\"">[ a-zA-Z]*</span>)*" + reg_all + "*</value>";
        public static string reg_author = @"<span class=""label"">作者: </span>" + reg_all + @"*等?\.?</div>";
        public static string reg_publisher = @"<span class=""FR_label"">出版商 </span><span>" + reg_all + "*</span>";
        public static string reg_issn = @"<span class=""FR_label"">ISSN: </span> <value>[0-9a-zA-Z\-]*</value>";
        public static string reg_jourName = @"<value>[\""\[\]\{\}\&\;\!\+\=\*\?\'\,\(\)\./:\- a-zA-Z0-9]*</value> </span>[&nbsp;]*<span class=""label"">[卷期特刊页出版年文献号丛书]*: </span><span class=""data_bold""> <value>[0-9a-zA-Z \-]*</value>";
        public static string reg_jourYear = @"</span>[ &nbsp;]*<span class=""label"">出版年: </span><span class=""data_bold""> <value>[0-9a-zA-Z \-]*</value>";
        public static string reg_useFre = @"<div class=""search-results-data-cite"">被引频次: (<a title=""查看引用本文的所有文献"" [\""\[\]\{\}\&\;\!\+\=\*\?\'\,\(\)\./:_\- a-zA-Z0-9]*>[0-9]*</a>)?([0-9]*<br>)?";
        public static string reg_use180 = "最近 180 天: \t\t    <span class=\"alum_text\">[0-9\\,]*</span>";
        public static string reg_use2013 = "2013 年至今: \t\t    <span class=\"alum_text\">[0-9\\,]*</span>";
        //public static string reg_eissn = @"<span class=""FR_label"">eISSN: </span> <value>[0-9a-zA-Z\-]*</value>";

        public static void GetParams(DataTable dt, string fileName) {
            string str = File.ReadAllText(fileName).Replace('\n', ' ').Replace('\r', ' ');
            string pageNo = Regex.Match(str, reg_PageNo).ToString();
            string pageCount = Regex.Match(str, reg_PageCount).ToString();

            //pageNo = pageNo.Substring(pageNo.IndexOf("value=\"") + 7, pageNo.IndexOf("\" size=") - pageNo.IndexOf("value=\"") - 7);
            //pageCount = pageCount.Substring(pageCount.IndexOf("bottom\">") + 8, pageCount.IndexOf("</span>") - pageCount.IndexOf("bottom\">") - 8).Replace(",", "");

            var titles = Regex.Matches(str, reg_title);
            string[] str_titles = new string[titles.Count];
            for (int i = 0; i < str_titles.Length; i++) {
                str_titles[i] = titles[i].ToString().Replace("<value lang_id=\"\">", "").Replace("<span class=\"hitHilite\">", "").Replace(@"</span>", "").Replace(@"</value>", "");
            }

            var authors = Regex.Matches(str, reg_author);
            string[] str_authors = new string[authors.Count];
            for (int i = 0; i < authors.Count; i++) {
                str_authors[i] = authors[i].ToString().Replace("<span class=\"label\">作者: </span>", "").Replace("</div>", "");
            }

            var publishers = Regex.Matches(str, reg_publisher);
            string[] str_publishers = new string[publishers.Count];
            for (int i = 0; i < publishers.Count; i++) {
                str_publishers[i] = publishers[i].ToString().Replace("<span class=\"FR_label\">出版商 </span><span>", "").Replace("</span>", "");
            }

            var issn = Regex.Matches(str, reg_issn);
            string[] str_issn = new string[issn.Count];
            for (int i = 0; i < issn.Count; i++) {
                str_issn[i] = issn[i].ToString().Replace("<span class=\"FR_label\">ISSN: </span> <value>", "").Replace("</value>", "");
            }

            // var jourName = Regex.Matches(str, reg_jourName);
            var jourYear = Regex.Matches(str, reg_jourYear);
            //string[] str_jourName = new string[jourName.Count];
            string[] str_jourYear = new string[jourYear.Count];
            for (int i = 0; i < jourYear.Count; i++) {
                //string[] tmp1 = jourName[i].ToString().Split('<');
                string[] tmp2 = jourYear[i].ToString().Split('<');
                //str_jourName[i] = tmp1[1].Substring(6, tmp1[1].Length - 6);
                str_jourYear[i] = tmp2[tmp2.Length - 2].Substring(6, tmp2[tmp2.Length - 2].Length - 6);
            }

            var useFre = Regex.Matches(str, reg_useFre);
            var use180 = Regex.Matches(str, reg_use180);
            var use2013 = Regex.Matches(str, reg_use2013);
            string[] str_useFre = new string[useFre.Count];
            string[] str_use180 = new string[use180.Count];
            string[] str_use2013 = new string[use2013.Count];

            for (int i = 0; i < useFre.Count; i++) {
                //str_useFre[i] = useFre[i].ToString().Replace("<div class=\"search-results-data-cite\">被引频次:", "").Replace("<br>", "");
                string[] tmp = useFre[i].ToString().Split('<');
                if (tmp.Length == 4) {
                    str_useFre[i] = tmp[2].Substring(tmp[2].IndexOf('>') + 1);
                } else {
                    str_useFre[i] = useFre[i].ToString().Replace("<div class=\"search-results-data-cite\">被引频次:", "").Replace("<br>", "");
                }

                str_use180[i] = use180[i].ToString().Replace("最近 180 天: \t\t    <span class=\"alum_text\">", "").Replace("</span>", "");
                str_use2013[i] = use2013[i].ToString().Replace("2013 年至今: \t\t    <span class=\"alum_text\">", "").Replace("</span>", "");
            }
            for (int i = 0; i < titles.Count; i++) {

                dt.Rows.Add(new string[] { str_titles[i], str_authors[i], str_publishers[i], str_jourYear[i], str_useFre[i], str_use180[i], str_use2013[i] });

                //string tmp = str_titles[i].Replace(',','_') + ","
                //    + str_authors[i].Replace(',', '_') + ","
                //    + str_publishers[i].Replace(',', '_') + ","
                //    //+ str_issn[i] + ","
                //    + str_jourName[i].Replace(',', '_') + ","
                //    + str_jourYear[i].Replace(',', '_') + ","
                //    + str_useFre[i].Replace(',', '_') + ","
                //    + str_use180[i].Replace(',', '_') + ","
                //    + str_use2013[i].Replace(',', '_');
                //resultList.Add(tmp);

            }

        }

        public void Analyse() {
            new Thread(new ThreadStart(GetAll)).Start();
        }

        private void GetAll() {
            dt.Columns.Add("标题");
            dt.Columns.Add("作者");
            dt.Columns.Add("出版商");
            //dt.Columns.Add("出版刊物名称");
            dt.Columns.Add("出版时间");
            dt.Columns.Add("被引频次");
            dt.Columns.Add("180天内使用次数");
            dt.Columns.Add("2013年以来使用次数");
            //list.Add("标题,作者,出版商,出版刊物名称,出版时间,被引频次,180天内使用次数,2013年以来使用次数");
            for (int i = 1; i <= count; i++) {
                HtmAnalyse.GetParams(dt, folderName + "\\result" + i + ".html");
                step = i;
            }
        }

        public bool SaveToExcel(DataTable dt, string fileName) {
            if (File.Exists(fileName)) {
                File.Delete(fileName);
            }
            File.Copy("model", fileName);
            string strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source =" + fileName + ";Extended Properties = Excel 8.0";
            OleDbConnection oleConnection = new OleDbConnection(strConnection);
            OleDbCommand command = new OleDbCommand();
            command.Connection = oleConnection;
            string sql = "";
            List<string> str_baks = new List<string>();
            try {
                oleConnection.Open();
                //int i = 0;

                foreach (DataRow dr in dt.Rows) {
                    try {
                        sql = string.Format("insert into [Sheet1$]([标题],[作者],[出版商],[出版时间],[被引频次],[180天内使用次数],[2013年以来使用次数]) values('{0}','{1}','{2}','{3}','{4}','{5}','{6}')",
                        dr[0].ToString().Replace("'", "''"), dr[1].ToString().Replace("'", "''"), dr[2].ToString().Replace("'", "''"), dr[3].ToString().Replace("'", "''"), dr[4].ToString().Replace("'", "''"), dr[5].ToString().Replace("'", "''"), dr[6].ToString().Replace("'", "''"));
                        //string sql = string.Format("insert into [Sheet1$] values(\"{0}\",\"{1}\")",
                        //    dr[0].ToString(), dr[1].ToString());
                        command.CommandText = sql;
                        command.ExecuteNonQuery();
                    } catch (Exception) {
                        string str_bak = string.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}", dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString(), dr[4].ToString(), dr[5].ToString(), dr[6].ToString());
                        str_baks.Add(str_bak);
                    }
                }
                if (str_baks.Count > 0) {
                    File.WriteAllLines(fileName + ".txt", str_baks.ToArray());
                }
                return true;
            } catch (Exception e) {
                throw e;
            } finally {
                oleConnection.Close();
            }




        }
    }
}
