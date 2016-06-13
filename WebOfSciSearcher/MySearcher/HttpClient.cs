using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MySearcher {
    public class HttpClient {
        Dictionary<string, string> uriParams = new Dictionary<string, string>();
        string jsesionid = "";
        HttpWebRequest request;
        HttpWebResponse response;
        string location;
        string body = "";
        byte[] btbody;
        string[] locationParam;
        string cookieUri;
        string resultHtm;
        string university;
        string searchYear = "2013-2014";

        public HttpClient(string university, string searchYear) {
            this.university = university.Replace(" ", "%20");
            this.searchYear = searchYear;
        }

        public bool OpenWeb() {
            string url = "http://www.webofknowledge.com";
            request = (HttpWebRequest)HttpWebRequest.Create(url);
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36";
            request.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8");
            request.Headers.Add("Accept-Encoding", "gzip,deflate,sdch");
            request.Host = "www.webofknowledge.com";
            request.KeepAlive = true;
            request.AllowAutoRedirect = false;
            response = (HttpWebResponse)request.GetResponse();

            //2
            OpenSetp2();
            //if (response.StatusCode == HttpStatusCode.Found) {
            //    return false;
            //    for (int i = 0; i < 10; i++) {
            //        OpenSetp2();
            //        if (response.StatusCode == HttpStatusCode.OK) break;
            //        if (i == 9) {
            //            return false;
            //        }
            //    }
            //}

            //3
            location = response.Headers.Get("Location");
            request = (HttpWebRequest)HttpWebRequest.Create(location);
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36";
            request.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8");
            request.Headers.Add("Accept-Encoding", "gzip,deflate,sdch");
            request.Host = "apps.webofknowledge.com";
            request.AllowAutoRedirect = false;
            request.CookieContainer = new CookieContainer();
            locationParam = response.Headers.Get("Location").ToString().Split('?')[1].Split('&');
            jsesionid = response.Cookies["JSESSIONID"].Value;
            foreach (var str in locationParam) {
                if (!uriParams.ContainsKey(str.Split('=')[0])) {
                    uriParams.Add(str.Split('=')[0], str.Split('=')[1]);
                }
                uriParams[str.Split('=')[0]] = str.Split('=')[1];
            }
            request.Headers.Add("Conection", "Keep-Alive");
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("SID", "\"" + uriParams["SID"] + "\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("CUSTOMER", "\"CAS National Sciences Library of Chinese Academy of Sciences\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("E_GROUP_NAME", "\"CAS Library of Beijing\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("JSESSIONID", jsesionid));
            response = (HttpWebResponse)request.GetResponse();
            string strs = new StreamReader(response.GetResponseStream()).ReadToEnd();


            //4
            location = response.Headers.Get("Location");
            request = (HttpWebRequest)HttpWebRequest.Create(location);
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36";
            request.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8");
            request.Headers.Add("Accept-Encoding", "gzip,deflate,sdch");
            request.Host = "apps.webofknowledge.com";
            request.Headers.Add("Conection", "Keep-Alive");
            request.AllowAutoRedirect = false;
            request.CookieContainer = new CookieContainer();
            locationParam = response.Headers.Get("Location").ToString().Split('?')[1].Split('&');
            jsesionid = response.Headers.Get("Set-Cookie").ToString().Split(';')[0].Split('=')[1];
            foreach (var str in locationParam) {
                if (!uriParams.ContainsKey(str.Split('=')[0])) {
                    uriParams.Add(str.Split('=')[0], str.Split('=')[1]);
                }
                uriParams[str.Split('=')[0]] = str.Split('=')[1];
            }
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("SID", "\"" + uriParams["SID"] + "\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("CUSTOMER", "\"CAS National Sciences Library of Chinese Academy of Sciences\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("E_GROUP_NAME", "\"CAS Library of Beijing\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("JSESSIONID", jsesionid));
            response = (HttpWebResponse)request.GetResponse();
            return response.StatusDescription == "OK";

        }

        private void OpenSetp2() {
            location = response.Headers.Get("Location");
            request = (HttpWebRequest)HttpWebRequest.Create(location);
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36";
            request.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8");
            request.Headers.Add("Accept-Encoding", "gzip,deflate,sdch");
            request.Host = "apps.webofknowledge.com";
            request.KeepAlive = true;
            request.AllowAutoRedirect = false;
            string cookestring = response.Headers.Get("Set-Cookie");
            request.CookieContainer = new CookieContainer();
            locationParam = response.Headers.Get("Location").ToString().Split('?')[1].Split('&');
            foreach (var str in locationParam) {
                if (!uriParams.ContainsKey(str.Split('=')[0])) {
                    uriParams.Add(str.Split('=')[0], str.Split('=')[1]);
                } else {
                    uriParams[str.Split('=')[0]] = str.Split('=')[1];
                }
            }
            cookieUri = "http://" + request.Host;
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("SID", "\"" + uriParams["SID"] + "\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("CUSTOMER", "\"CAS National Sciences Library of Chinese Academy of Sciences\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("E_GROUP_NAME", "\"CAS Library of Beijing\""));
            response = (HttpWebResponse)request.GetResponse();
        }

        public void Search(string i) {
            string type = "OG";
            if (i == "1") {
                type = "FO";
            } else if (i == "2") {
                type = "AD";
            }
            location = "http://apps.webofknowledge.com/WOS_GeneralSearch.do";
            request = (HttpWebRequest)HttpWebRequest.Create(location);
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.Referer = "http://apps.webofknowledge.com/WOS_GeneralSearch_input.do?product=WOS&search_mode=GeneralSearch&SID=" + uriParams["SID"] + "&preferencesSaved=";
            request.Headers.Add("Origin", "http://apps.webofknowledge.com");
            request.Accept = "*/*";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36";
            request.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8");
            request.Headers.Add("Accept-Encoding", "gzip,deflate,sdch");
            request.Host = "apps.webofknowledge.com";
            request.Headers.Add("Conection", "Keep-Alive");
            request.Headers.Add("Cache-Control", "max-age=0");
            request.AllowAutoRedirect = false;
            request.CookieContainer = new CookieContainer();
            jsesionid = response.Headers.Get("Set-Cookie").ToString().Split(';')[0].Split('=')[1];
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("SID", "\"" + uriParams["SID"] + "\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("CUSTOMER", "\"CAS National Sciences Library of Chinese Academy of Sciences\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("E_GROUP_NAME", "\"CAS Library of Beijing\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("JSESSIONID", jsesionid));


            //       body = "fieldCount=3&action=search&product=WOS&search_mode=GeneralSearch&SID=" +
            //uriParams["SID"] +
            //"&max_field_count=25&max_field_notice=%E6%B3%A8%E6%84%8F%3A%20%E6%97%A0%E6%B3%95%E6%B7%BB%E5%8A%A0%E5%8F%A6%E4%B8%80%E5%AD%97%E6%AE%B5%E3%80%82&input_invalid_notice=%E6%A3%80%E7%B4%A2%E9%94%99%E8%AF%AF%3A%20%E8%AF%B7%E8%BE%93%E5%85%A5%E6%A3%80%E7%B4%A2%E8%AF%8D%E3%80%82" +
            //"&exp_notice=%E6%A3%80%E7%B4%A2%E9%94%99%E8%AF%AF%3A%20%E4%B8%93%E5%88%A9%E6%A3%80%E7%B4%A2%E8%AF%8D%E5%8F%AF%E5%9C%A8%E5%A4%9A%E4%B8%AA%E5%AE%B6%E6%97%8F%E4%B8%AD%E6%89%BE%E5%88%B0%20(&+" +
            //"input_invalid_notice_limits=%20%3Cbr%2F%3E%E6%B3%A8%3A%20%E6%BB%9A%E5%8A%A8%E6%A1%86%E4%B8%AD%E6%98%BE%E7%A4%BA%E7%9A%84%E5%AD%97%E6%AE%B5%E5%BF%85%E9%A1%BB%E8%87%B3%E5%B0%91%E4%B8%8E%E4%B8%80%E4%B8%AA%E5%85%B6%E4%BB%96%E6%A3%80%E7%B4%A2%E5%AD%97%E6%AE%B5%E7%9B%B8%E7%BB%84%E9%85%8D%E3%80%82" +
            //"&sa_params=WOS%7C%7C" + uriParams["SID"] + "%7Chttp%3A%2F%2Fapps.webofknowledge.com%7C'&formUpdated=true&" +
            //"value(input1)=" + university + "&value(select1)="+type+"&value(hidInput1)=&" + "value(bool_1_2)=AND&" +
            //"value(input2)=" + searchYear + "&value(select2)=PY&value(hidInput2)=&" + "value(bool_2_3)=AND&" +
            //"value(input3)=Article&value(select3)=DT&value(hidInput3)=&" +
            //"limitStatus=collapsed&ss_lemmatization=On&ss_spellchecking=Suggest&SinceLastVisit_UTC=&SinceLastVisit_DATE=&period=Range%20Selection&range=ALL&startYear=1864&endYear=2016&update_back2search_link_param=yes&ssStatus=display%3Anone&ss_showsuggestions=ON&ss_query_language=auto&ss_numDefaultGeneralSearchFields=1&rs_sort_by=PY.D%3BLD.D%3BSO.A%3BVL.D%3BPG.A%3BAU.A&";

            //查询文献与综述  单位是地址
            body = "fieldCount=6&action=search&product=WOS&search_mode=GeneralSearch&SID=" +
     uriParams["SID"] +
     "&max_field_count=25&max_field_notice=%E6%B3%A8%E6%84%8F%3A%20%E6%97%A0%E6%B3%95%E6%B7%BB%E5%8A%A0%E5%8F%A6%E4%B8%80%E5%AD%97%E6%AE%B5%E3%80%82&input_invalid_notice=%E6%A3%80%E7%B4%A2%E9%94%99%E8%AF%AF%3A%20%E8%AF%B7%E8%BE%93%E5%85%A5%E6%A3%80%E7%B4%A2%E8%AF%8D%E3%80%82" +
     "&exp_notice=%E6%A3%80%E7%B4%A2%E9%94%99%E8%AF%AF%3A%20%E4%B8%93%E5%88%A9%E6%A3%80%E7%B4%A2%E8%AF%8D%E5%8F%AF%E5%9C%A8%E5%A4%9A%E4%B8%AA%E5%AE%B6%E6%97%8F%E4%B8%AD%E6%89%BE%E5%88%B0%20(&+" +
     "input_invalid_notice_limits=%20%3Cbr%2F%3E%E6%B3%A8%3A%20%E6%BB%9A%E5%8A%A8%E6%A1%86%E4%B8%AD%E6%98%BE%E7%A4%BA%E7%9A%84%E5%AD%97%E6%AE%B5%E5%BF%85%E9%A1%BB%E8%87%B3%E5%B0%91%E4%B8%8E%E4%B8%80%E4%B8%AA%E5%85%B6%E4%BB%96%E6%A3%80%E7%B4%A2%E5%AD%97%E6%AE%B5%E7%9B%B8%E7%BB%84%E9%85%8D%E3%80%82" +
     "&sa_params=WOS%7C%7C" + uriParams["SID"] + "%7Chttp%3A%2F%2Fapps.webofknowledge.com%7C'&formUpdated=true&" +
     "value(input1)=" + university + "&value(select1)=" + type + "&value(hidInput1)=&" + "value(bool_1_2)=AND&" +
     "value(input2)=" + searchYear + "&value(select2)=PY&value(hidInput2)=&" + "value(bool_2_3)=AND&" +
     "value(input3)=Article&value(select3)=DT&value(hidInput3)=&" + "value(bool_3_4)=OR&" +
      "value(input4)=" + university + "&value(select4)=" + type + "&value(hidInput4)=&" + "value(bool_4_5)=AND&" +
     "value(input5)=" + searchYear + "&value(select5)=PY&value(hidInput5)=&" + "value(bool_5_6)=AND&" +
     "value(input6)=Review&value(select6)=DT&value(hidInput6)=&" +
     "limitStatus=collapsed&ss_lemmatization=On&ss_spellchecking=Suggest&SinceLastVisit_UTC=&SinceLastVisit_DATE=&period=Range%20Selection&range=ALL&startYear=1864&endYear=2016&update_back2search_link_param=yes&ssStatus=display%3Anone&ss_showsuggestions=ON&ss_query_language=auto&ss_numDefaultGeneralSearchFields=1&rs_sort_by=PY.D%3BLD.D%3BSO.A%3BVL.D%3BPG.A%3BAU.A&";


            btbody = Encoding.UTF8.GetBytes(body);
            request.ContentLength = btbody.Length;
            request.GetRequestStream().WriteTimeout = 4000;
            request.GetRequestStream().Write(btbody, 0, btbody.Length);
            request.GetRequestStream().Close();
            response = (HttpWebResponse)request.GetResponse();

            //if (response.StatusDescription == "Moved Temporarily") {
            //    OpenSetp2();
            //}

            //get
            location = response.Headers.Get("Location");
            request = (HttpWebRequest)HttpWebRequest.Create(location);
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36";
            request.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8");
            request.Headers.Add("Accept-Encoding", "gzip,deflate,sdch");
            request.Host = "apps.webofknowledge.com";
            request.Headers.Add("Cache-Control", "max-age=0");
            request.AllowAutoRedirect = false;
            request.CookieContainer = new CookieContainer();
            request.Referer = "http://apps.webofknowledge.com/WOS_GeneralSearch_input.do?product=WOS&search_mode=GeneralSearch&SID=" + uriParams["SID"] + "&preferencesSaved=";

            locationParam = response.Headers.Get("Location").ToString().Split('?')[1].Split('&');
            jsesionid = response.Cookies["JSESSIONID"].Value;
            foreach (var str in locationParam) {
                if (!uriParams.ContainsKey(str.Split('=')[0])) {
                    uriParams.Add(str.Split('=')[0], str.Split('=')[1]);
                }
                uriParams[str.Split('=')[0]] = str.Split('=')[1];
            }
            request.Headers.Add("Conection", "Keep-Alive");
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("SID", "\"" + uriParams["SID"] + "\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("CUSTOMER", "\"CAS National Sciences Library of Chinese Academy of Sciences\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("E_GROUP_NAME", "\"CAS Library of Beijing\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("JSESSIONID", jsesionid));
            response = (HttpWebResponse)request.GetResponse();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns>返回总页数</returns>
        public int ChangeNum50() {
            int count = 0;
            location = "http://apps.webofknowledge.com/summary.do?product=WOS&parentProduct=WOS&search_mode=GeneralSearch&qid=1&SID=" + uriParams["SID"] + "&page=1&action=changePageSize&pageSize=50";
            request = (HttpWebRequest)HttpWebRequest.Create(location);
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36";
            request.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8");
            request.Headers.Add("Accept-Encoding", "gzip,deflate,sdch");
            request.Host = "apps.webofknowledge.com";
            request.AllowAutoRedirect = false;
            request.CookieContainer = new CookieContainer();
            request.Referer = response.ResponseUri.ToString();

            jsesionid = response.Cookies["JSESSIONID"].Value;

            request.Headers.Add("Conection", "Keep-Alive");
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("SID", "\"" + uriParams["SID"] + "\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("CUSTOMER", "\"CAS National Sciences Library of Chinese Academy of Sciences\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("E_GROUP_NAME", "\"CAS Library of Beijing\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("JSESSIONID", jsesionid));
            response = (HttpWebResponse)request.GetResponse();

            string str = new StreamReader(response.GetResponseStream()).ReadToEnd();
            string reg_count = "<span id=\"pageCount.bottom\">[\\,\\d]*</span>";
            reg_count = System.Text.RegularExpressions.Regex.Match(str, reg_count).ToString();
            reg_count = reg_count.Substring(reg_count.IndexOf("bottom\">") + 8, reg_count.IndexOf("</span>") - reg_count.IndexOf("bottom\">") - 8).Replace(",", "");
            int.TryParse(reg_count, out count);
            return count;
        }

        public void ChangePage(int i) {
            location = "http://apps.webofknowledge.com/summary.do?product=WOS&parentProduct=WOS&search_mode=GeneralSearch&parentQid=&qid=1&SID=" + uriParams["SID"] + "&&update_back2search_link_param=yes&page=" + i;
            request = (HttpWebRequest)HttpWebRequest.Create(location);
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.63 Safari/537.36";
            request.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8");
            request.Headers.Add("Accept-Encoding", "gzip,deflate,sdch");
            request.Host = "apps.webofknowledge.com";
            request.AllowAutoRedirect = false;
            request.CookieContainer = new CookieContainer();
            request.Referer = response.ResponseUri.ToString();

            //locationParam = response.Headers.Get("Location").ToString().Split('?')[1].Split('&');
            try {
                jsesionid = response.Cookies["JSESSIONID"].Value;
            } catch (Exception) {
            }

            //foreach (var str in locationParam) {
            //    if (!uriParams.ContainsKey(str.Split('=')[0])) {
            //        uriParams.Add(str.Split('=')[0], str.Split('=')[1]);
            //    }
            //    uriParams[str.Split('=')[0]] = str.Split('=')[1];
            //}
            request.Headers.Add("Conection", "Keep-Alive");
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("SID", "\"" + uriParams["SID"] + "\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("CUSTOMER", "\"CAS National Sciences Library of Chinese Academy of Sciences\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("E_GROUP_NAME", "\"CAS Library of Beijing\""));
            request.CookieContainer.Add(new Uri(cookieUri), new Cookie("JSESSIONID", jsesionid));
            response = (HttpWebResponse)request.GetResponse();
        }

        public string GetResultHtm() {
            System.IO.StreamReader searchReader = new System.IO.StreamReader(response.GetResponseStream());
            resultHtm = searchReader.ReadToEnd();
            return resultHtm;
        }
    }
}
