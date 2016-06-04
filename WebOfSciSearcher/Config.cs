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
        public float matchGate = 0.6f;
        public string httpWeb;

    }
}
