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
                    using (OleDbDataAdapter adpater = new OleDbDataAdapter("select * from schoolConfig", conn)) {
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
                    config.titleRegex = row["titleRegex"].ToString().Trim();
                    config.titleLeft = int.Parse(row["titleLeft"].ToString());
                    config.titleRight = int.Parse(row["titleRight"].ToString());
                    config.contentReplace = row["contentReplace"].ToString().Trim();
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

            List<string> strArray1 = str1.ToLower().Split(' ').ToList();
            List<string> strArray2 = str2.ToLower().Split(' ').ToList();


            foreach (var item in strArray1) {
                if (strArray2.Contains(item)) {
                    strArray2.Remove(item);
                    matchValue++;
                    continue;
                }
            }

            strArray1 = str1.ToLower().Split(' ').ToList();
            strArray2 = str2.ToLower().Split(' ').ToList();
            foreach (var item in strArray2) {
                if (strArray1.Contains(item)) {
                    strArray1.Remove(item);
                    matchValue++;
                    continue;
                }
            }

            strArray1 = str1.ToLower().Split(' ').ToList();
            strArray2 = str2.ToLower().Split(' ').ToList();
            matchValue = matchValue / (strArray1.Count + strArray2.Count);

            return Math.Round(matchValue, 2);
        }

    }
}
