using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebOfSciSearcher {
    public class Utils {

        static string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=wos.mdb";


        /// <summary>
        /// 读取配置文件
        /// </summary>
        /// <returns></returns>
        public static DataTable GetMdbConfig() {
            DataTable dtReturn = new DataTable();
            try {
                using (OleDbConnection conn = new OleDbConnection(connStr)) {
                    conn.Open();
                    using (OleDbDataAdapter adpater = new OleDbDataAdapter("select * from schoolConfig order by schoolIndex", conn)) {
                        adpater.Fill(dtReturn);
                    }
                }
            } catch (Exception) {
                return null;
            }
            return dtReturn;
        }


        /// <summary>
        /// 获取配置
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, Config> GetConfig() {
            DataTable dt = GetMdbConfig();
            Dictionary<string, Config> dicReturn = new Dictionary<string, Config>();
            if (dt != null) {
                foreach (DataRow row in dt.Rows) {
                    Config config = new Config();
                    config.schoolIndex = int.Parse(row["schoolIndex"].ToString());
                    config.schoolName = row["schoolName"].ToString().Trim();
                    config.url1 = row["url1"].ToString().Trim();
                    config.url2 = row["url2"].ToString().Trim();
                    config.titleRegex = row["titleRegex"].ToString().Trim();
                    config.titleLeft = int.Parse(row["titleLeft"].ToString());
                    config.titleRight = int.Parse(row["titleRight"].ToString());
                    string[] replaces = row["contentReplace"].ToString().Split('|');
                    config.contentReplace = new List<string[]>();
                    for (int i = 0; i < replaces.Length; i++) {
                        if (!string.IsNullOrEmpty(replaces[0])) {
                            config.contentReplace.Add(new string[] { replaces[i].Split(',')[0], replaces[i].Split(',')[1].Replace("空格", " ") });
                        }
                    }

                    replaces = row["urlReplace"].ToString().Split('|');
                    config.urlReplace = new List<string[]>();
                    for (int i = 0; i < replaces.Length; i++) {
                        if (!string.IsNullOrEmpty(replaces[0])) {
                            config.urlReplace.Add(new string[] { replaces[i].Split(',')[0], replaces[i].Split(',')[1].Replace("空格", " ") });
                        }
                    }

                    config.matchGate = float.Parse(row["matchGate"].ToString());
                    config.httpWeb = row["httpWeb"].ToString();
                    config.articleRegex = row["articleRegex"].ToString();
                    config.downloadRegex = row["downloadRegex"].ToString();
                    string tmp = row["downloadIndex"].ToString();
                    int.TryParse(tmp, out config.downloadIndex);
                    config.detailWeb = row["detailWeb"].ToString();

                    config.dateAssionRegex = row["dateAssionRegex"].ToString();
                    tmp = row["dateAssionIndex"].ToString();
                    int.TryParse(tmp, out config.dateAssionIndex);


                    config.dateAvailableRegex = row["dateAvailableRegex"].ToString();
                    tmp = row["dateAvailableIndex"].ToString();
                    int.TryParse(tmp, out config.dateAvailableIndex);

                    config.dateIssuRegex = row["dateIssuRegex"].ToString();
                    tmp = row["dateIssuIndex"].ToString();
                    int.TryParse(tmp, out config.dateIssuIndex);

                    config.languageRegex = row["languageRegex"].ToString();
                    tmp = row["languageIndex"].ToString();
                    int.TryParse(tmp, out config.languageIndex);

                    config.rightsRegex = row["rightsRegex"].ToString();
                    tmp = row["rightsIndex"].ToString();
                    int.TryParse(tmp, out config.rightsIndex);

                    config.rightsUriRegex = row["rightsUriRegex"].ToString();
                    tmp = row["rightsUriIndex"].ToString();
                    int.TryParse(tmp, out config.rightsUriIndex);

                    config.typeRegex = row["typeRegex"].ToString();
                    tmp = row["typeIndex"].ToString();
                    int.TryParse(tmp, out config.typeIndex);

                    dicReturn.Add(config.schoolName, config);
                }
            }
            return dicReturn;
        }

        /// <summary>
        /// 计算匹配度
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <returns></returns>
        public static double MatchValue(string str1, string str2) {
            double matchValue = 0;

            List<string> strArray1 = ToList(str1);
            List<string> strArray2 = ToList(str2);

            foreach (var item in strArray1) {
                if (strArray2.Contains(item)) {
                    strArray2.Remove(item);
                    matchValue++;
                    continue;
                }
            }

            strArray1 = ToList(str1);
            strArray2 = ToList(str2);
            foreach (var item in strArray2) {
                if (strArray1.Contains(item)) {
                    strArray1.Remove(item);
                    matchValue++;
                    continue;
                }
            }

            strArray1 = ToList(str1);
            strArray2 = ToList(str2);
            matchValue = matchValue / (strArray1.Count + strArray2.Count);

            return Math.Round(matchValue, 2);
        }

        private static List<string> ToList(string str1) {
            List<string> lReturn = str1.ToLower().Split(' ').ToList();
            int count = lReturn.Count;
            for (int i = 0; i < count; i++) {
                lReturn.Remove("");
            }
            return lReturn;
        }

    }
}
