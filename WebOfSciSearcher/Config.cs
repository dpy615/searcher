using System.Collections.Generic;

namespace WebOfSciSearcher {
    public class Config {
        public int schoolIndex;
        public string schoolName;
        public string url1;
        public string url2;
        public string titleRegex;
        public int titleLeft;
        public int titleRight;
        public List<string[]> contentReplace;
        public string articleRegex;
        public string downloadRegex;
        public int downloadIndex;

        public float matchGate = 0.6f;
        public string httpWeb;

        public string detailWeb;

        public string dateAssionRegex;
        public int dateAssionIndex;

        public string dateAvailableRegex;
        public int dateAvailableIndex;

        public string dateIssuRegex;
        public int dateIssuIndex;

        public string languageRegex;
        public int languageIndex;

        public string rightsRegex;
        public int rightsIndex;

        public string rightsUriRegex;
        public int rightsUriIndex;

        public string typeRegex;
        public int typeIndex;

    }
}
